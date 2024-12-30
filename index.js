const XlsxPopulate = require('xlsx-populate');
const dialog = require('node-file-dialog');
const path = require('path');


//prepared functions
let _ren = undefined;
function fetchByVin(vin){
	console.log(`${arguments.callee.name}('${vin}');`);
	return fetchRen().then(ren=>{
		//console.log(ren);
		return fetch(`https://dataovozidlech.cz/api/Vozidlo/GetVehicleInfo?vin=${vin}`, {
			headers: {
				'_ren': ren,
				'Referer': 'https://dataovozidlech.cz/vyhledavani',
				//'Referrer-Policy': 'strict-origin-when-cross-origin',
				//'accept': 'application/json, text/plain, */*',
				//'accept-language': 'cs-CZ,cs;q=0.9,en;q=0.8',
				//'sec-fetch-dest': 'empty',
				//'sec-fetch-mode': 'cors',
				//'sec-fetch-site': 'same-origin',
				//'cookie': 'BIGipServer~md-dk~POOL-MD-DK-P-443=rd2966o00000000000000000000ffff0ae68fa5o443',
			},
			body: null
		}).then(res=>res.json()).then(data=>{
			if(data.message){
				console.log(data.message);
				throw Error(data.message);
			}
			return data;
		})
	});
}

function fetchRen(){
	if(_ren){return Promise.resolve(_ren)}
	else{
		console.log('fetching ren');
		return fetch('https://dataovozidlech.cz/').then(res=>{
			_ren = res.headers.get('_ren');
			return _ren;
		});
	}
}


//code execution
dialog({type:'open-file'}).then(filePaths => {
	console.log(filePaths);

	filePaths.forEach(filePath=>{
		const parsedPath = path.parse(filePath);

		XlsxPopulate.fromFileAsync(filePath).then(wb=>{
			//wb - workbook
			let vins = [], vin = true;
			const promises = [];
			const sheet = wb.sheet('flotila');
			for(let i = 0; i<500 && vin; i++){
				vin = sheet.cell(`C${i+3}`).value();
				if(!vin)break;
				vins[i] = vin;
				promises.push(new Promise(resolve=>{setTimeout(()=>{resolve((()=>{return fetchByVin(vins[i]);})())}, (60000/10)*0*i)}));
			}

			return Promise.allSettled(promises).then(vehicleDatas=>{
				console.log("allSettled");
				vehicleDatas.forEach((prom, i)=>{
					//console.log(prom);
					const vehicle = prom.value;
					console.log(i, prom.status, (vehicle||[]).find(x=>x.name=='TovarniZnacka')?.value);
					if(prom.status == 'rejected') throw Error(prom.reason);
					[
						{name: 'CisloOrv', col:'D'},
						{name: 'CisloTp', col:'E'},
						{name: 'TovarniZnacka', col:'F'},
						{name: 'Typ', col:'H'},
						{name: 'ObchodniOznaceni', col:'G'},
						{name: '', col:'I'}, //???
						{name: 'DatumPrvniRegistrace', col:'J'},
						{name: 'DatumPrvniRegistraceVCr', col:'K'},

						{name: 'VozidloDruh', col:'M'},

						{name: 'MotorMaxVykon', col:'P', fn: x=>{return Number(x.split('/')[0]?.trim())||0}},
						{name: 'MotorZdvihObjem', col:'Q'},
						{name: 'Palivo', col:'R', fn: x=>{return x + ({'NM':' / Diesel', 'BA': ' / Benzin'}[x]??'')}},
						{name: 'HmotnostiPripPov', col:'S', fn: x=>{return Number(x.split('/')[0]?.trim())||0}},

						{name: 'VozidloKaroserieMist', col:'U', fn: x=>{return Number(x.split('/')[1]?.trim())||0}},
						{name: 'VozidloKaroserieMist', col:'V', fn: x=>{return Number(x.split('/')[0]?.trim())||0}},
					].forEach(att=>{
						const fn = att.fn ?? (x=>{return x});
						sheet.cell(`${att.col}${i+3}`).value(fn(vehicle.find(x=>x.name==att.name)?.value)) //Velký TP
					});
				});

				const outFilePath = `${parsedPath.dir}/${parsedPath.name}-vyplneno${parsedPath.ext}`;
				console.log(outFilePath);
				return wb.toFileAsync(outFilePath);
			});
		}).catch(err=>{
			console.log('CHYBA!:');
			console.log(err.message);
		});
	});
});