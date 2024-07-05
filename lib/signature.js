const crypto = require("crypto");

function signature(algorithm, data, privateKey) {
    try {
        const sign = crypto.createSign(algorithm);
        sign.update(Buffer.from(data));
        const result = sign.sign(privateKey, 'base64');
        return result;;
    } catch (error) {
        return null
    }
}

module.exports = {
    signature
};