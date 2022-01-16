const { exportKeychain } = require("../lib")

const [, , filePath] = process.argv

exportKeychain(filePath)
