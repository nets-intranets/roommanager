require("dotenv").config();
const path = require("path")

var log4js = require("log4js");
var logger = log4js.getLogger();

log4js.configure({
    appenders: {
      everything: { type: 'file', keepFileExt :true,filename: path.join(__dirname,'../../logs.log'), maxLogSize: 10485760, backups: 3, compress: true },
      out: { type: 'stdout', layout: { type: 'basic' } } 
    
    },
    categories: {
      default: { appenders: [ 'everything','out' ], level: 'debug'}
    }
    });
    
    
    
    
    
module.exports.logger = logger    