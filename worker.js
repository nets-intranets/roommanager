// https://tomasz.janczuk.org/2013/07/application-initialization-of-nodejs.html
var log = require("./helpers/log")
var pack = require("./package.json")
var running = true
var POLLINTERVAL = 5000
var woody = require("./jobs/batch-all")

const MILLISECSBETWEENUPDATES =  1000 * 60 * 60
log.logger.info(pack.name,pack.version)
log.logger.info("Booting...")


    async function loop(){
        log.logger.info("Refreshing...")
        try {
           var result = await  woody.run()
        } catch (error) {
            log.logger.error(error.message,error)
        }finally
        {
            setTimeout(loop, MILLISECSBETWEENUPDATES)
        }

    return 
    }

loop()

