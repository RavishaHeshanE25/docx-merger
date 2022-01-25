var XMLSerializer = require("xmldom").XMLSerializer;
var DOMParser = require("xmldom").DOMParser;

var prepareMediaFilesAsync = async function (files, media) {
  var count = 1;

  await files.forEach(async function (zip, index) {
    await zip.folder("word/media").forEach(async function (relativePath, _file) {
      if (/^word\/media/.test(relativePath) && relativePath.length > 11) {
        media[count] = {};
        media[count].oldTarget = relativePath;
        media[count].newTarget = relativePath
          .replace(/[0-9]/, "_" + count)
          .replace("word/", "");
        media[count].fileIndex = index;
        await updateMediaRelationsAsync(zip, count, media);
        await updateMediaContentAsync(zip, count, media);
        count++;
      }
    });
  });
};

var updateMediaRelationsAsync = async function (zip, count, _media) {
  var xmlString = await zip.file("word/_rels/document.xml.rels").async("text");
  var xml = new DOMParser().parseFromString(xmlString, "text/xml");

  var childNodes = xml.getElementsByTagName("Relationships")[0].childNodes;
  var serializer = new XMLSerializer();

  for (var node in childNodes) {
    if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
      var target = childNodes[node].getAttribute("Target");
      if ("word/" + target == _media[count].oldTarget) {
        _media[count].oldRelID = childNodes[node].getAttribute("Id");

        childNodes[node].setAttribute("Target", _media[count].newTarget);
        childNodes[node].setAttribute(
          "Id",
          _media[count].oldRelID + "_" + count
        );
      }
    }
  }

  var startIndex = xmlString.indexOf("<Relationships");
  xmlString = xmlString.replace(
    xmlString.slice(startIndex),
    serializer.serializeToString(xml.documentElement)
  );

  zip.file("word/_rels/document.xml.rels", xmlString);
};

var updateMediaContentAsync = async function (zip, count, _media) {
  var xmlString = await zip.file("word/document.xml").async("text");

  xmlString = xmlString.replace(
    new RegExp(_media[count].oldRelID + '"', "g"),
    _media[count].oldRelID + "_" + count + '"'
  );

  zip.file("word/document.xml", xmlString);
};

var copyMediaFilesAsync = async function (base, _media, _files) {
  for (var media in _media) {
    var content = await _files[_media[media].fileIndex]
      .file(_media[media].oldTarget)
      .async("uint8array");

    base.file("word/" + _media[media].newTarget, content);
  }
};

module.exports = {
  prepareMediaFilesAsync,
  updateMediaRelationsAsync,
  updateMediaContentAsync,
  copyMediaFilesAsync,
};
