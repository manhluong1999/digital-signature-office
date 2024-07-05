const { DOMParser, XMLSerializer } = require('@xmldom/xmldom');

function removeSelfClosingTags(xml) {
    const split = xml.split("/>");
    console.log('split',split);
    let newXml = "";
    for (let i = 0; i < split.length - 1; i++) {
        const edsplit = split[i].split("<");
        console.log('edsplit',edsplit);
        newXml += split[i] + "></" + edsplit[edsplit.length - 1].split(" ")[0] + ">";
    }
    return newXml + split[split.length - 1];
}
function orderAttributes(node) {
    const attributes = Array.from(node.attributes);
    const listAttribute = attributes.map((item) => item.name)
    listAttribute.sort()
    const orderedAttributes = listAttribute.map(attrName => attributes.find(attr => attr.name === attrName));
    orderedAttributes.forEach(attr => node.removeAttribute(attr.name));
    orderedAttributes.forEach(attr => node.setAttribute(attr.name, attr.value));
}
function applyRelationshipTransform(xmlString, listId = []) {
    const doc = new DOMParser().parseFromString(xmlString, 'application/xml');
    const relationships = xpath.select("//*[local-name(.)='Relationship']", doc);
    relationships.sort((a, b) => {
        const idA = a.getAttribute('Id');
        const idB = b.getAttribute('Id');
        return idA.localeCompare(idB);
    });
    const xml = new XMLSerializer()
    const sortedDoc = new DOMParser().parseFromString('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>', 'application/xml');
    const sortedRoot = sortedDoc.documentElement;
    relationships.forEach(item => {
        const idNode = item.getAttribute('Id');
        if (listId.includes(idNode)) {
            const importedNode = sortedDoc.importNode(item, true);
            importedNode.setAttribute('TargetMode', 'Internal')
            orderAttributes(importedNode);
            sortedRoot.appendChild(importedNode);
        }
    });
    return removeSelfClosingTags(xml.serializeToString(sortedRoot));
}

function canonicalizationXml(algorithm) {

}

module.exports = {
    applyRelationshipTransform
};