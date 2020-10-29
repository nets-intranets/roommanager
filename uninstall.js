// https://github.com/coreybutler/node-windows

var Service = require('node-windows').Service;
var path = require("path")

// Create a new service object
var svc = new Service({
  name:'Room Manager',
  description: 'Room Manager service.',
  script:  path.join(__dirname,"worker.js")
});

// Listen for the "install" event, which indicates the
// process is available as a service.
svc.on('uninstall',function(){
  console.log('Uninstall complete.');
  console.log('The service exists: ',svc.exists);
});

svc.uninstall()