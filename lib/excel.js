const fs = require('fs')
const JSZip = require('jszip');
const path = require('path');
const { SignedXml } = require('xml-crypto');
const { getHashDigest } = require('./hash');
const xpath = require('xpath')
const updateContentTypeXml = async (zip) => {
    const contentTypePath = '[Content_Types].xml'
    const contentTypesXml = await zip.file(contentTypePath).async('text')

    const data = parseStringToXmlObj(contentTypesXml)

    const numOfSignature = data.Types.Override.filter((item) => item['$'].ContentType == 'application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml' && item['$'].PartName.includes('/_xmlsignatures/sig')).length + 1

    const doc = new DOMParser().parseFromString(contentTypesXml, 'application/xml');

    const newElement = doc.createElement('Override');
    newElement.setAttribute('PartName', `/_xmlsignatures/sig${numOfSignature}.xml`)
    newElement.setAttribute('ContentType', 'application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml')

    const parentNode = doc.getElementsByTagName('Types')[0];
    parentNode.appendChild(newElement);
    const updatedContentTypesXml = new XMLSerializer().serializeToString(doc);

    zip.file(contentTypePath, updatedContentTypesXml);
    return numOfSignature
}
const updatRelsXml = async (zip) => {
    const filePath = '_rels/.rels'
    const xmlString = await zip.file(filePath).async('text')
    const doc = new DOMParser().parseFromString(xmlString, 'application/xml');

    const newElement = doc.createElement('Relationship');
    newElement.setAttribute('Id', `rId4`)
    newElement.setAttribute('Type', 'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin')
    newElement.setAttribute('Target', `_xmlsignatures/origin.sigs`)

    const parentNode = doc.getElementsByTagName('Relationships')[0];
    parentNode.appendChild(newElement);
    const updatedXml = new XMLSerializer().serializeToString(doc);

    zip.file(filePath, updatedXml);
}
const updatOriginSigsRelsXml = async (zip, numOfSignature) => {
    const filePath = '_xmlsignatures/_rels/origin.sigs.rels'
    const xml = await zip.file(filePath).async('text')

    const doc = new DOMParser().parseFromString(xml, 'application/xml');
    const newElement = doc.createElement('Relationship');
    newElement.setAttribute('Id', `rId${numOfSignature}`)
    newElement.setAttribute('Type', 'http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature')
    newElement.setAttribute('Target', `sig${numOfSignature}.xml`)

    const parentNode = doc.getElementsByTagName('Relationships')[0];
    parentNode.appendChild(newElement);
    const updatedXml = new XMLSerializer().serializeToString(doc);

    zip.file(filePath, updatedXml);
}
const insertSigXml = async (zip, numOfSignature, privateKeyPem, certData) => {
    let sigXml = fs.readFileSync(path.resolve(__dirname, '../templates/sig_template.xml'), 'utf8')

    //digest value _rels/.rels
    const data = {
        signature_time: new Date().toISOString()
    }
    const relsPath = '_rels/.rels'
    const relsXmlStr = await zip.file(relsPath).async('text')
    const relsTransformed = applyRelationshipTransform(relsXmlStr, ['rId1']);
    data.digest_value_rels = getHashDigest(relsTransformed)

    // workbookRels
    const workbookRelsPath = 'xl/_rels/workbook.xml.rels'
    const workbookRelsXmlStr = await zip.file(workbookRelsPath).async('text')
    const workbookRelsTransformed = applyRelationshipTransform(workbookRelsXmlStr, ['rId1', 'rId2', 'rId3', 'rId4']);
    data.digest_value_workbook_rels = getHashDigest(workbookRelsTransformed)

    //printerSettings1
    const printerSettings1Path = 'xl/printerSettings/printerSettings1.bin'
    const printerSettings1File = await zip.file(printerSettings1Path).async('binaryString')
    data.digest_value_printer_settings1 = getHashDigest(printerSettings1File, 'binary')

    // sharedStrings
    const sharedStringsPath = 'xl/sharedStrings.xml'
    const sharedStringsXmlStr = await zip.file(sharedStringsPath).async('text')
    data.digest_value_shared_string = getHashDigest(sharedStringsXmlStr)

    // styles
    const stylesPath = 'xl/styles.xml'
    const stylesXmlStr = await zip.file(stylesPath).async('text')
    data.digest_value_styles = getHashDigest(stylesXmlStr)

    // theme1
    const theme1Path = 'xl/theme/theme1.xml'
    const theme1XmlStr = await zip.file(theme1Path).async('text')
    data.digest_value_theme1 = getHashDigest(theme1XmlStr)

    // workbook
    const workbookPath = 'xl/workbook.xml'
    const workbookXmlStr = await zip.file(workbookPath).async('text')
    data.digest_value_workbook = getHashDigest(workbookXmlStr)

    // sheet1Rels
    const sheet1RelsPath = 'xl/worksheets/_rels/sheet1.xml.rels'
    const sheet1RelsXmlStr = await zip.file(sheet1RelsPath).async('text')
    const sheet1RelsTransformed = applyRelationshipTransform(sheet1RelsXmlStr, ['rId1']);
    data.digest_value_sheet1_rels = getHashDigest(sheet1RelsTransformed)

    // sheet1
    const sheet1Path = 'xl/worksheets/sheet1.xml'
    const sheet1XmlStr = await zip.file(sheet1Path).async('text')
    data.digest_value_sheet1 = getHashDigest(sheet1XmlStr)

    //certDigest
    data.digest_value_cert = getHashDigest(Buffer.from(certData.certificate, 'base64'), 'utf8')
    data.certificate = certData.certificate
    const issuerAttributes = certData.cert.issuer.attributes;
    const issuer = {};
    issuerAttributes.forEach(attr => {
        issuer[attr.shortName] = attr.value;
    });
    data.x509_issuer = issuerAttributes.map((item) => `${item.shortName}=${item.value}`).join(', ')
    data.x509_serial_number = BigInt(`0x${certData.cert.serialNumber}`).toString(10);
    for (const key of Object.keys(data)) {
        const reg = new RegExp(`{${key}}`, 'g')
        sigXml = sigXml.replace(reg, data[key])
    }

    const sig = new SignedXml({
        privateKey: privateKeyPem,
        publicCert: certData.certificatePem,
        canonicalizationAlgorithm: 'http://www.w3.org/TR/2001/REC-xml-c14n-20010315',
        signatureAlgorithm: 'http://www.w3.org/2001/04/xmldsig-more#rsa-sha256',
    });

    const doc = new DOMParser().parseFromString(sigXml);

    const packageObjectElement = xpath.select(`//*[@Id="idPackageObject"]`, doc)[0];
    const canonicalizedXml1 = sig.getCanonXml(["http://www.w3.org/TR/2001/REC-xml-c14n-20010315"], packageObjectElement);
    data.digest_value_package_object = getHashDigest(canonicalizedXml1);

    const officeObjectElement = xpath.select(`//*[@Id="idOfficeObject"]`, doc)[0];
    const canonicalizedXml2 = sig.getCanonXml(["http://www.w3.org/TR/2001/REC-xml-c14n-20010315"], officeObjectElement);
    data.digest_value_office_object = getHashDigest(canonicalizedXml2);

    const signedPropertiesElement = xpath.select(`//*[@Id="idSignedProperties"]`, doc)[0];
    let canonicalizedXml3 = sig.getCanonXml(["http://www.w3.org/TR/2001/REC-xml-c14n-20010315"], signedPropertiesElement);

    const reg = new RegExp(' xmlns="http://www.w3.org/2000/09/xmldsig#"', 'g')
    canonicalizedXml3 = canonicalizedXml3.replace(reg, '')

    const reg1 = new RegExp('xmlns:xd="http://uri.etsi.org/01903/v1.3.2#"', 'g')
    canonicalizedXml3 = canonicalizedXml3.replace(reg1, 'xmlns="http://www.w3.org/2000/09/xmldsig#" xmlns:xd="http://uri.etsi.org/01903/v1.3.2#"')

    data.digest_value_signed_properties = getHashDigest(canonicalizedXml3);

    sigXml = sigXml.replace(`{digest_value_package_object}`, data.digest_value_package_object)
    sigXml = sigXml.replace(`{digest_value_office_object}`, data.digest_value_office_object)
    sigXml = sigXml.replace(`{digest_value_signed_properties}`, data.digest_value_signed_properties)
   
    const docFull = new DOMParser().parseFromString(sigXml);
    const signedInfoNode = xpath.select(`//*[local-name(.)='SignedInfo']`, docFull)[0];
    const signedInfoXml = sig.getCanonXml(["http://www.w3.org/TR/2001/REC-xml-c14n-20010315"], signedInfoNode);

    fs.writeFileSync(`test/signed_info${numOfSignature}.xml`, Buffer.from(signedInfoXml))

    data.signature_value = await signature('RSA-SHA256',signedInfoXml, privateKeyPem, false)

    sigXml = sigXml.replace(`{signature_value}`, data.signature_value)

    zip.file(`_xmlsignatures/sig${numOfSignature}.xml`, sigXml);
    console.log('data', data);
    fs.writeFileSync(`test/sig${numOfSignature}.xml`, Buffer.from(sigXml))
}

async function addDigitalSignature(buffer, privateKeyPath, certPath, capass = '12345678') {
    const zip = new JSZip();
    await zip.loadAsync(buffer);

    const numOfSignature = await updateContentTypeXml(zip)

    if (numOfSignature == 1) {
        await updatRelsXml(zip)
        zip.folder('_xmlsignatures/_rels');
        zip.file('_xmlsignatures/origin.sigs', '');

        const sigRelsXml = fs.readFileSync(path.resolve(__dirname, '../templates/sigs_rels_template.xml'), 'utf8')
        zip.file('_xmlsignatures/_rels/origin.sigs.rels', sigRelsXml);
    }

    const certData = getCertFromCer(certPath, true)
    const privateKey = getPrivateKeyFromP12(privateKeyPath, capass, true)

    await updatOriginSigsRelsXml(zip, numOfSignature)

    await insertSigXml(zip, numOfSignature, privateKey, certData)

    const signedExcelData = await zip.generateAsync({ type: 'nodebuffer' });

    return signedExcelData
}

module.exports = {
    addDigitalSignature
};

