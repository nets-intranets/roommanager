require("dotenv").config();
const logger = require("../../helpers/log").logger;
const _ = require("lodash");
const path = require("path");
const fs = require("fs-extra");
const json = require("format-json");
const helpers = require("../../helpers");

function run() {
  return new Promise(async (resolve, reject) => {
    console.log("Building Room Master data");
    var filepath = path.join(
      __dirname,
      "../../../data",
      process.env.AADDOMAIN,
      "rooms-sharepoint.json"
    );
    var roomsText = fs.readFileSync(filepath, "utf8").replace(/^\uFEFF/, "");
    var rooms = JSON.parse(roomsText);

    var keys = _.keys(rooms);
    var masterdata = [];

    keys.forEach((key) => {
      var item = rooms[key];
      var room = item.room ? item.room : {}
      var building = item.building ? item.building : {}
      var location = item.location ? item.location : {}
      

      masterdata.push({
        primarySMTPAddress: helpers.keepNull(room.Title),
        displayName: helpers.keepNull(room.Display_x0020_Name),
        provisioning_x0020_Status: helpers.keepNull(room.Provisioning_x0020_Status),
        capacity: helpers.keepNull(room.Capacity),
        audioDeviceName: helpers.keepNull(room.AudioDeviceName),
        videoDeviceName: helpers.keepNull(room.VideoDeviceName),
        displayDeviceName: helpers.keepNull(room.DisplayDeviceName),
        isWheelChairAccessible: room.IsWheelChairAccessible === "Yes" ,
        phone: helpers.keepNull(room.Phone),
        floor: helpers.keepNull(room.Floor),
        floorLabel: helpers.keepNull(room.FloorLabel),
        building : helpers.keepNull(building.Title),
        geoCoordinates : building.GeoCoordinates ? helpers.keepNull(building.GeoCoordinates) : "0;0",
        street:  helpers.keepNull(location.Street),
        city:  helpers.keepNull(location.City),
        state:  helpers.keepNull(location.State),
        postalCode:  helpers.keepNull(location.PostalCode),
        countryOrRegion:  helpers.keepNull(location.CountryOrRegion)



      });
    });

    var filepathOut = path.join(
      __dirname,
      "../../../data",
      process.env.AADDOMAIN,
      "rooms-masterdata.json"
    );
    var rooms = fs.writeFileSync(filepathOut, json.plain(masterdata));

    console.log("Done building Room Master data");
    resolve();
  });
}

module.exports.run = run;

//run();
