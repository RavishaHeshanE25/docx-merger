var JSZip = require("jszip");

var Style = require("./merge-styles");
var Media = require("./merge-media");
var RelContentType = require("./merge-relations-and-content-type");
var bulletsNumbering = require("./merge-bullets-numberings");

function DocxMerger(options = {}) {
  this._body = [];
  this._header = [];
  this._footer = [];
  this._Basestyle = options.style || "source";
  this._style = [];
  this._numbering = [];
  this._pageBreak =
    typeof options.pageBreak !== "undefined" ? !!options.pageBreak : true;
  this._files = [];
  this._contentTypes = {};

  this._media = {};
  this._rel = {};

  this._builder = this._body;

  var _self = this;
  this.addFilesAsync = async function (files = []) {
    await Promise.all(files.map(async function (file) {
      _self._files.push(await new JSZip().loadAsync(file));
    }));
    if (_self._files.length > 0) {
      await _self.mergeBodyAsync(_self._files);
    }
  };

  this.insertPageBreak = function () {
    var pb =
      '<w:p> \
            <w:r> \
                <w:br w:type="page"/> \
            </w:r> \
            </w:p>';

    _self._builder.push(pb);
  };

  this.insertRaw = function (xml) {
    _self._builder.push(xml);
  };

  this.mergeBodyAsync = async function (files = []) {
    _self._builder = _self._body;

    await RelContentType.mergeContentTypesAsync(files, _self._contentTypes);
    await Media.prepareMediaFilesAsync(files, _self._media);
    await RelContentType.mergeRelationsAsync(files, _self._rel);

    await bulletsNumbering.prepareNumberingAsync(files);
    await bulletsNumbering.mergeNumberingAsync(files, _self._numbering);

    await Style.prepareStylesAsync(files, _self._style);
    await Style.mergeStylesAsync(files, _self._style);

    await Promise.all(files.map(async function (zip, index) {
      var xml = await zip.file("word/document.xml").async("text");
      xml = xml.substring(xml.indexOf("<w:body>") + 8);
      xml = xml.substring(0, xml.indexOf("</w:body>"));
      xml = xml.substring(0, xml.lastIndexOf("<w:sectPr"));
  
      _self.insertRaw(xml);
      if (_self._pageBreak && index < files.length - 1) _self.insertPageBreak();
    }));
  };

  this.saveAsync = async function (type) {
    var zip = _self._files[0];
    var xml = await zip.file("word/document.xml").async("text");
    var startIndex = xml.indexOf("<w:body>") + 8;
    var endIndex = xml.lastIndexOf("<w:sectPr");
    
    xml = xml.replace(xml.slice(startIndex, endIndex), _self._body.join(""));

    await RelContentType.generateContentTypesAsync(zip, _self._contentTypes);
    await Media.copyMediaFilesAsync(zip, _self._media, _self._files);
    await RelContentType.generateRelationsAsync(zip, _self._rel);
    await bulletsNumbering.generateNumberingAsync(zip, _self._numbering);
    await Style.generateStylesAsync(zip, _self._style);

    zip.file("word/document.xml", xml);

    return await zip.generateAsync({
      type,
      compression: "DEFLATE",
      compressionOptions: {
        level: 4,
      },
    });
  };
}

module.exports = DocxMerger;
