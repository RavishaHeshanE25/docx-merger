var XMLSerializer = require("xmldom").XMLSerializer;
var DOMParser = require("xmldom").DOMParser;

var prepareNumberingAsync = async function (files) {
  var serializer = new XMLSerializer();

  await Promise.all(
    files.map(async function (zip, index) {
      var xmlBin = zip.file("word/numbering.xml");
      if (!xmlBin) {
        return;
      }
      var xmlString = await xmlBin.async("text");
      var xml = new DOMParser().parseFromString(xmlString, "text/xml");
      var nodes = xml.getElementsByTagName("w:abstractNum");

      for (var node in nodes) {
        if (/^\d+$/.test(node) && nodes[node].getAttribute) {
          var absID = nodes[node].getAttribute("w:abstractNumId");
          nodes[node].setAttribute("w:abstractNumId", absID);
        }
      }

      var numNodes = xml.getElementsByTagName("w:num");

      for (var node in numNodes) {
        if (/^\d+$/.test(node) && numNodes[node].getAttribute) {
          var ID = numNodes[node].getAttribute("w:numId");
          numNodes[node].setAttribute("w:numId", ID);
          var absrefID = numNodes[node].getElementsByTagName("w:abstractNumId");
          for (var i in absrefID) {
            if (absrefID[i].getAttribute) {
              var iId = absrefID[i].getAttribute("w:val");
            }
          }
        }
      }

      var startIndex = xmlString.indexOf("<w:numbering ");
      xmlString = xmlString.replace(
        xmlString.slice(startIndex),
        serializer.serializeToString(xml.documentElement)
      );

      zip.file("word/numbering.xml", xmlString);
    })
  );
};

var mergeNumberingAsync = async function (files, _numbering) {
  await Promise.all(
    files.map(async function (zip) {
      var xmlBin = zip.file("word/numbering.xml");
      if (!xmlBin) {
        return;
      }
      var xml = await xmlBin.async("text");

      xml = xml.substring(
        xml.indexOf("<w:abstractNum "),
        xml.indexOf("</w:numbering")
      );

      _numbering.push(xml);
    })
  );
};

var generateNumberingAsync = async function (zip, _numbering) {
  var xmlBin = zip.file("word/numbering.xml");
  if (!xmlBin) {
    return;
  }
  var xml = await xmlBin.async("text");
  var startIndex = xml.indexOf("<w:abstractNum ");
  var endIndex = xml.indexOf("</w:numbering>");

  if (startIndex !== -1 && endIndex !== -1) {
    xml = xml.replace(xml.slice(startIndex, endIndex), _numbering.join(""));
  } else {
    xml = `${xml.substring(0, xml.length - 2)}>${_numbering.join(
      ""
    )}</w:numbering>`;
  }

  zip.file("word/numbering.xml", xml);
};

module.exports = {
  prepareNumberingAsync,
  mergeNumberingAsync,
  generateNumberingAsync,
};
