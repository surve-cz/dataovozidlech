const { contextBridge } = require("electron");
const fetchProcess = require('index.js').runProcess;

contextBridge.exposeInMainWorld("versions", {
    node: () => process.versions.node,
    chrome: () => process.versions.chrome,
    electron: () => process.versions.electron,
    // we can also expose variables, not just functions
});

contextBridge.exposeInMainWorld('backend', {
	process: (file)=> fetchProcess(file)
});