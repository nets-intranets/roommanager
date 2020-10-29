require("dotenv").config();
const _ = require("lodash");
const path = require("path");
const fs = require("fs-extra");
const json = require("format-json");

const helpers = require("../../helpers");
async function run(){
    return new Promise(async (resolve, reject) => {
        const KEY = "primarySMTPAddress";
        const FIELDS = ["displayName"];
        var master = require("../../../data/"+process.env.AADDOMAIN+"/rooms-masterdata.json")
        var replica = require("../../../data/"+process.env.AADDOMAIN+"/rooms-slavedata.json")
        var diffs = await helpers.compareArrays(
          master,
          replica,
          KEY,
          FIELDS
        );
        var filepathOut = path.join(
            __dirname,
            "../../../data",
            process.env.AADDOMAIN,
            "rooms-diffs.json"
          );
          fs.writeFileSync(filepathOut, json.plain(diffs));
      
      
        return resolve(diffs)
                
    });


}

module.exports.run = run;
run()