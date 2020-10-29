require("dotenv").config();
var _ = require("lodash");
var request = require("request");
var https = require("https");
const logger = require("../log").logger;

var auth = {};
var config = {
  clientId: process.env.ROOMMGRAPPCLIENT_ID,
  clientSecret: process.env.ROOMMGRAPPCLIENT_SECRET,
  tokenEndpoint:
    "https://login.microsoftonline.com/" +
    process.env.ROOMMGRAPPCLIENT_DOMAIN +
    "/oauth2/token",
};

var graph = {};

function extractEmails(text) {
  return text.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/gi);
}

graph.readGroupMembers = function (token, groupId, cb) {
  get(
    token,
    "https://graph.microsoft.com/beta/groups/" +
      groupId +
      "/members?$select=accountEnabled,userPrincipalName,id",
    cb
  );
};

graph.readGroupMembers2 = function (token, groupId) {
  return get2(
    token,
    "https://graph.microsoft.com/v1.0/groups/" +
      groupId +
      "/members?$select=accountEnabled,userPrincipalName,id"
  );
};

graph.readGroupMemberShipFor = function (token, upn, cb) {
  get(token, "https://graph.microsoft.com/v1.0/users/" + upn + "/memberOf", cb);
};

graph.readManager = function (token, userPrincipalName, index, cb) {
  get(
    token,
    "https://graph.microsoft.com/v1.0/users/" + userPrincipalName + "/manager",
    cb,
    null,
    index
  );
};

graph.readUser = function (token, userPrincipalName, index, cb) {
  get(
    token,
    "https://graph.microsoft.com/v1.0/users/" + userPrincipalName,
    cb,
    null,
    index
  );
};

graph.readUser2 = function (token, userPrincipalName) {
  return get2(
    token,
    "https://graph.microsoft.com/v1.0/users/" + userPrincipalName
  );
};

graph.readUsers = function (token, index, cb) {
  get(token, "https://graph.microsoft.com/beta/users/", cb, null, index);
};

graph.readUsers2 = function (token) {
  return get2(token, "https://graph.microsoft.com/beta/users/");
};

graph.readContacts = function (token, index, cb) {
  get(token, "https://graph.microsoft.com/beta/contacts/", cb, null, index);
};

graph.readContacts2 = function (token) {
  return get2(token, "https://graph.microsoft.com/beta/contacts/");
};

graph.readGuests2 = function (token) {
  return get2(token, "https://graph.microsoft.com/beta/contacts/");
};

graph.readGroups = function (token, index, cb) {
  get(token, "https://graph.microsoft.com/beta/groups/", cb, null, index);
};

graph.readGroups2 = function (token, index, cb) {
  return get2(token, "https://graph.microsoft.com/beta/groups/");
};
graph.readUserPhoto = function (token, userPrincipalName, index, cb) {
  getFile(
    token,
    "graph.microsoft.com",
    "/v1.0/users/" + userPrincipalName + "/photo/$value",
    cb,
    null,
    index
  );
};

graph.matchingGroups = function (token, searchText, cb) {
  get(
    token,
    `https://graph.microsoft.com/v1.0/groups?$select=displayName,id,mail,mailNickname,groupTypes,description&$filter=startsWith(displayName,'${searchText}')`,
    cb
  );
};

graph.getGroupDocuments = function (token, groupId, folderName, cb) {
  get(
    token,
    `https://graph.microsoft.com/beta/groups/${groupId}/drive/items/root:${folderName}:/children`,
    cb
  );
};

graph.getGroupDrive = function (token, groupId, cb) {
  get(token, `https://graph.microsoft.com/beta/groups/${groupId}/drive`, cb);
};

graph.apppointments = function (token, upn, cb) {
  get(
    token,
    `https://graph.microsoft.com/v1.0/users/${upn}/events?$select=subject,organizer,attendees,start,end,location&$filter=end lt DateTime('2020-12-31T09:13:28')`,
    cb
  );
};

graph.getGroupSite = function (token, groupId, cb) {
  get(
    token,
    `https://graph.microsoft.com/beta/groups/${groupId}/sites/root`,
    cb
  );
};

graph.getDriveChildren = function (token, driveId, itemId, cb) {
  get(
    token,
    `https://graph.microsoft.com/beta/drives/${driveId}/items/${itemId}/children`,
    cb
  );
};

graph.getGroupRootFolders = function (token, groupId, cb) {
  get(
    token,
    `https://graph.microsoft.com/beta/groups/${groupId}/drive/items/root/children`,
    cb
  );
};

graph.getSharePointListItems2 = function (
  token,
  hostName,
  siteCollectionName,
  listName
) {
  //  https://graph.microsoft.com/v1.0/sites/christianiabpos.sharepoint.com:/sites/intranets-location:/lists/SAP Nets Work Addresses/items
  return get2(
    token,
    `https://graph.microsoft.com/v1.0/sites/${hostName}:/sites/${siteCollectionName}:/lists/${listName}/items?expand=fields`
  );
};

function lookupFile(token, driveId, id, name, iterations, cb) {
  // Content have been copied
  graph.getDriveChildren(token, driveId, id, function (err, result) {
    if (err) {
      return cb(err);
    }
    // content of destination folder

    for (let index = 0; index < result.length; index++) {
      const file = result[index];

      console.log("Comparing", file.name, name);
      if (file.name === name) {
        // match found
        console.log("Match found");
        return cb(null, file);
      }
    }

    function retry() {
      console.log("Retrying", iterations);
      //
      if (iteration > 10) {
        return cb(
          `Think that the file has been created, but could not find a reference to if. Filename is ${name} - Tryed ${iterations} to find it `
        );
      }
      lookupFile(token, driveId, id, name, iterations + 1, cb);
    }
    // no match found so we will wait a second a try again
    setTimeout(retry, 1000);
  });
}

graph.createDistributionList = function (
  token,
  mailNickname,
  displayName,
  description
) {
  var payload = {
    displayName,
    description,
    groupTypes: ["Unified"],
    mailEnabled: true,
    mailNickname,

    securityEnabled: false,
  };
  return post2(token, "https://graph.microsoft.com/v1.0/groups", payload);
};

function post2(token, url, payload) {
  return new Promise((resolve, reject) => {
    post(token, url, payload, (error, result) => {
      if (error) {
        return reject(error);
      }
      resolve(result);
    });
  });
}

graph.newDoc = function (token, sourceGroupId, itemId, driveId, id, name, cb) {
  post(
    token,
    `https://graph.microsoft.com/beta/groups/${sourceGroupId}/drive/items/${itemId}/copy`,
    {
      parentReference: {
        driveId: driveId,
        id: id,
      },
      name: name,
    },
    function (err) {
      // no result from that call - by design from Microsoft
      if (err) {
        return cb(err);
      }

      // so we fill try to find a referece our self
      lookupFile(token, driveId, id, name, 0, cb);
    }
  );
};

//config.clientSecret = process.env[config.clientId]

auth.getAccessToken = function () {
  var retrycount = 0;
  return new Promise((resolve, reject) => {
    var requestParams = {
      grant_type: "client_credentials",
      client_id: config.clientId,
      client_secret: config.clientSecret,
      resource: "https://graph.microsoft.com",
    };
    function doPost() {
      request.post(
        {
          url: config.tokenEndpoint,
          form: requestParams,
        },
        function (err, response, body) {
          try {
            var parsedBody = JSON.parse(body);

            if (err) {
              logger.warn("auth.getAccessToken ", err);
              reject(err);
            } else if (parsedBody.error) {
              logger.warn("auth.getAccessToken ", parsedBody.error_description);
              reject(parsedBody.error_description);
            } else {
              resolve(parsedBody.access_token);
            }
          } catch (error) {
            function sleep(ms) {
              console.log("sleeping", ms);
              return new Promise((resolve, reject) => {
                setTimeout(() => {
                  console.log("woke up");
                  resolve();
                }, ms);
              });
            }
            async function handleError() {
              //logger.warn("auth.getAccessToken ", error);
              retrycount++;
              if (retrycount > 3) reject("auth timed out");
              await sleep(2000);
              doPost();
            }
            handleError();
          }
        }
      );
    }

    doPost();
  });
};

function post(token, url, body, cb) {
  request.post(
    {
      url: url,
      headers: {
        "content-type": "application/json",
        authorization: "Bearer " + token,
      },
      body: JSON.stringify(body),
    },
    function (err, response, body) {
      var parsedBody;

      if (err) {
        return cb(err, response);
      } else {
        parsedBody = body === "" ? {} : JSON.parse(body);
        if (parsedBody.error) {
          return cb(parsedBody.error);
        } else {
          cb(null, parsedBody);
        }
      }
    }
  );
}

function sleep(ms) {
  return new Promise((resolve, reject) => {
    setTimeout(() => {
      resolve();
    }, ms);
  });
}

async function get(token, url, cb, appendTo, index, retrycount) {
  if (process.stdout.clearLine) {
    process.stdout.clearLine();
    process.stdout.cursorTo(0);
    console.log(`Processing ${url}`);
  }

  console.log("getting", url, appendTo === true);
  request.get(
    {
      url: url,
      headers: {
        "content-type": "application/json",
        authorization: "Bearer " + token,
      },
    },
    function (err, response, body) {
      var parsedBody;

      if (err) {
        var errCounter = retrycount ? retrycount + 1 : 1;
        if (errCounter > 3) {
          logger.warn("Error, retry max.passed getting ", url, err);
          return cb(err, null, index);
        } else {
          // retry
          console.log("------------------------");
          console.log("getting", url, appendTo === true);
          console.log("retry #", errCounter);
          console.log("------------------------");
          return get(token, url, cb, appendTo, index, errCounter);
        }
      }

      if (response.headers["content-type"].indexOf("json") === -1) {
        return cb(null, body, index, response.headers["content-type"]);
      }
      parsedBody = JSON.parse(body);
      if (parsedBody.error) {
        return cb(parsedBody.error, null, index);
      }

      if (parsedBody["@odata.nextLink"]) {
        //console.log("Sleeping")
        var data;
        if (!appendTo) {
          data = parsedBody.value;
        } else {
          data = appendTo.concat(parsedBody.value);
        }
        setTimeout(() => {
          //console.log("Waking")

          get(token, parsedBody["@odata.nextLink"], cb, data, index);
        }, 200);
        return;
      }
      if (process.stdout.clearLine) {
        process.stdout.clearLine();
        process.stdout.cursorTo(0);
      }
      if (appendTo) {
        cb(null, appendTo.concat(parsedBody.value), index);
      } else {
        if (parsedBody.value) {
          cb(null, parsedBody.value, index);
        } else {
          cb(null, parsedBody, index);
        }
      }
    }
  );
}
function get2(token, url, cb) {
  return new Promise((resolve, reject) => {
    get(token, url, (error, result) => {
      if (error) {
        return reject(error);
      }
      resolve(result);
    });
  });
}

function getFile(token, host, path, cb, appendTo, index, retrycount) {
  var options = {
    host, //: 'graph.microsoft.com',
    path, // : "/v1.0/users/ngjoh@nets.eu/photo/$value",
    method: "GET",
    headers: {
      Authorization: "Bearer " + token,
    },
  };

  https
    .get(options, function (response) {
      response.setEncoding("binary"); /* This is very very necessary! */
      var body = "";
      response.on("data", function (d) {
        body += d;
      });
      response.on("end", function () {
        var error;
        if (response.statusCode === 200) {
          /* save as "normal image" */
          // fs.writeFile('./public/img/image.jpeg', body, 'binary',  function(err){
          //     if (err) throw err
          //     console.log('Image saved ok')
          // })
          /* callback - for example show in template as base64 image */
          cb(null, body, index); //new Buffer(body, 'binary').toString('base64'),index);
        } else {
          error = new Error();
          error.code = response.statusCode;
          error.message = response.statusMessage;
          // The error body sometimes includes an empty space
          // before the first character, remove it or it causes an error.
          body = body.trim();
          //error.innerError = JSON.parse(body).error;
          cb(error, null);
        }
      });
    })
    .on("error", function (e) {
      cb(e, null);
    });
}

function parallel(tasks, maxParallelTasks, run, onSuccess, onError) {
  if (!run || !_.isFunction(run)) {
    logger.fatal("No run method");
    process.exit(1);
  }

  if (!tasks || !_.isArray(tasks)) {
    logger.fatal("Tasks is not an array");
    process.exit(1);
  }

  if (!maxParallelTasks || !_.isInteger(maxParallelTasks)) {
    process.exit(1);
  }

  var currentTasks = 0;
  var toRun = _.cloneDeep(tasks);

  return new Promise((resolve, reject) => {
    function takeNext() {
      console.log("Tick",toRun.length,currentTasks);
      if (toRun.length === 0 && currentTasks === 0) {
        resolve();
        return
      }

      while ((currentTasks < maxParallelTasks) && (toRun.length > 0)) {
        console.log("Spanning new task");
        var task = toRun.pop();
        currentTasks++;

        run(task)
          .then((result) => {
            if (onSuccess) {
              onSuccess(task,result);
            }
            currentTasks--;
          })
          .catch((error) => {
            if (onError) {
              onError(task, error);
            }
            currentTasks--;
          });
      }

      setTimeout(takeNext, 100);
    }
    takeNext();
  });
}

module.exports = {
  auth,
  graph,
  parallel,
};
