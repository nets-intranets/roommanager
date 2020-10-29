const path = require("path");
const fs = require("fs-extra");
const _ = require("lodash");
const { forEach } = require("lodash");
/**
 * Ensure that the necessary structures for temporary storage is present
 */
function ensureDirs() {
  fs.ensureDirSync(path.join(__dirname, "../../data"));
}

function getNew(master, replica, key) {
  return new Promise((resolve, reject) => {
    if (!master) reject("No master");
    if (!replica) reject("No replica");

    var newItems = [];
    var replicaDictionary = _.groupBy(replica, key);
    master.forEach((element) => {
      if (!replicaDictionary[element[key]]) {
        newItems.push(element);
      }
    });

    resolve(newItems);
  });
}

function getDeleted(master, replica, key) {
  return new Promise((resolve, reject) => {
    if (!master) reject("No master");
    if (!replica) reject("No replica");

    var deletedItems = [];
    var masterDictionary = _.groupBy(master, key);
    replica.forEach((element) => {
      if (!masterDictionary[element[key]]) {
        deletedItems.push(element);
      }
    });

    resolve(deletedItems);
  });
}

function getChanged(master, replica, key, fields) {
  return new Promise((resolve, reject) => {
    if (!master) reject("No master");
    if (!replica) reject("No replica");
    var changedItems = [];
    var replicaDictionary = _.groupBy(replica, key);
    master.forEach((masterElement) => {
      var hasChanged = false;
      var replicaElement = replicaDictionary[masterElement[key]];
      if (replicaElement) {
        $changes = [];
        fields.forEach((field) => {
          if (masterElement[field] !== replicaElement[0][field]) {
            hasChanged = true;
            $changes.push({
              field,
              master: masterElement[field],
              replica: replicaElement[0][field],
            });
          }
        });

        if (hasChanged) {
          masterElement.$changes = $changes;
          changedItems.push(masterElement);
        }
      }
    });

    resolve(changedItems);
  });
}

function compareArrays(master, replica, key, fields) {
  return new Promise(async (resolve, reject) => {
    var newItems = await getNew(master, replica, key).catch((error) => {
      console.log("error compareArrays, newItems", error);
    });
    var deletedItems = await getDeleted(master, replica, key).catch((error) => {
      console.log("error compareArrays, newItems", error);
    });
    var changedItems = await getChanged(master, replica, key, fields).catch(
      (error) => {
        console.log("error compareArrays, newItems", error);
      }
    );
    resolve({ newItems, deletedItems, changedItems });
  });
}

function compareMemberships(master, replica) {
  return new Promise((resolve, reject) => {
    var masterDictionary = {};
    master.forEach((m) => {
      masterDictionary[m.alias + ":" + m.upn] = true;
    });

    var replicaDictionary = {};
    replica.forEach((m) => {
      replicaDictionary[m.alias + ":" + m.upn] = true;
    });

    var toAdd = [];
    var toRemove = [];

    master.forEach((m) => {
      var key = m.alias + ":" + m.upn;

      var replica = replicaDictionary[key];
      if (!replica) toAdd.push(m);
    });

    replica.forEach((m) => {
      var key = m.alias + ":" + m.upn;
      var master = masterDictionary[key];
      if (!master) toRemove.push(m);
    });

    resolve({ toAdd, toRemove });
  });
}

function keepNull(v){
  return v?v:"$$null$$"
}


module.exports = {
  keepNull,
  ensureDirs,
  getNew,
  getDeleted,
  getDeleted,
  compareArrays,
  compareMemberships,
};
