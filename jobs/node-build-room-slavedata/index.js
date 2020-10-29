require("dotenv").config();
const logger = require("../../helpers/log").logger;
const _ = require("lodash");
const path = require("path");
const fs = require("fs-extra");
const json = require("format-json");
const helpers = require("../../helpers");
function run() {
  return new Promise(async (resolve, reject) => {
    console.log("Building Room Slave data");
    var filepath = path.join(
      __dirname,
      "../../../data",
      process.env.AADDOMAIN,
      "rooms-exchange.json"
    );
    var roomsText = fs.readFileSync(filepath, "utf8").replace(/^\uFEFF/, "");
    var rooms = JSON.parse(roomsText);

    var keys = _.keys(rooms);
    var slaveData = [];

    keys.forEach((key) => {

       var item = rooms[key];
       var mailbox = item.mailbox ? item.mailbox : {}
       var place = item.place ? item.place : {}
      // var location = item.location ? item.location : {}
      

      slaveData.push({
        primarySMTPAddress: helpers.keepNull(mailbox.PrimarySmtpAddress),
        displayName: helpers.keepNull(mailbox.DisplayName),
        // provisioning_x0020_Status: room.Provisioning_x0020_Status,
         capacity: helpers.keepNull(place.Capacity),
         audioDeviceName: helpers.keepNull(place.AudioDeviceName),
         videoDeviceName: helpers.keepNull(place.VideoDeviceName),
         displayDeviceName: helpers.keepNull(place.DisplayDeviceName),
         isWheelChairAccessible: place.IsWheelChairAccessible ? "Yes": "No",
         phone: helpers.keepNull(place.Phone),
         floor: helpers.keepNull(place.Floor),
         floorLabel: helpers.keepNull(place.FloorLabel),
         building : helpers.keepNull(place.Building),
          geoCoordinates : helpers.keepNull(place.GeoCoordinates),
         street:  helpers.keepNull(place.Street),
         city:  helpers.keepNull(place.City),
         state:  helpers.keepNull(place.State),
         postalCode:  helpers.keepNull(place.PostalCode),
         countryOrRegion:  helpers.keepNull(place.CountryOrRegion)



      });
    });

    var filepathOut = path.join(
      __dirname,
      "../../../data",
      process.env.AADDOMAIN,
      "rooms-slavedata.json"
    );
    var rooms = fs.writeFileSync(filepathOut, json.plain(slaveData));

    console.log("Done building Room Slave data");
    resolve();
  });
}

module.exports.run = run;


