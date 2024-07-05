function getHashDigest(data, encoding = 'utf8', algo = 'sha256') {
    const shasum = crypto.createHash(algo);
    shasum.update(data, encoding);
    const res = shasum.digest("base64");
    return res;
}
module.exports = {
    getHashDigest
};