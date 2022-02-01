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

  this.addFilesAsync = async function (files = []) {
    var self = this;
    var zippedFiles = [];
    if (self._files.length > 0) {
      self.insertPageBreak();
    }
    await Promise.all(
      files.map(async function (file) {
        var zippedFile = await new JSZip().loadAsync(file);
        zippedFiles.push(zippedFile);
        self._files.push(zippedFile);
      })
    );
    if (zippedFiles.length > 0) {
      await self.mergeBodyAsync(zippedFiles);
    }
  };

  this.insertPageBreak = function () {
    var pb =
      '<w:p> \
        <w:r> \
          <w:br w:type="page"/> \
        </w:r> \
      </w:p>';

    this._builder.push(pb);
  };

  this.insertRaw = function (xml) {
    this._builder.push(xml);
  };

  this.mergeBodyAsync = async function (files = []) {
    var self = this;
    this._builder = this._body;

    await RelContentType.mergeContentTypesAsync(files, this._contentTypes);
    await Media.prepareMediaFilesAsync(files, this._media);
    await RelContentType.mergeRelationsAsync(files, this._rel);

    await bulletsNumbering.prepareNumberingAsync(files);
    await bulletsNumbering.mergeNumberingAsync(files, this._numbering);

    await Style.prepareStylesAsync(files, this._style);
    await Style.mergeStylesAsync(files, this._style);

    await Promise.all(
      files.map(async function (zip, index) {
        var xml = await zip.file("word/document.xml").async("text");
        xml = xml.substring(xml.indexOf("<w:body>") + 8);
        xml = xml.substring(0, xml.indexOf("</w:body>"));
        xml = xml.substring(0, xml.lastIndexOf("<w:sectPr"));

        self.insertRaw(xml);
        if (self._pageBreak && index < files.length - 1) self.insertPageBreak();
      })
    );
  };

  this.saveAsync = async function (type) {
    var zip = this._files[0];

    var xml = await zip.file("word/document.xml").async("text");
    var startIndex = xml.indexOf("<w:body>") + 8;
    var endIndex = xml.lastIndexOf("<w:sectPr");

    xml = xml.replace(xml.slice(startIndex, endIndex), this._body.join(""));

    await RelContentType.generateContentTypesAsync(zip, this._contentTypes);
    await Media.copyMediaFilesAsync(zip, this._media, this._files);
    await RelContentType.generateRelationsAsync(zip, this._rel);
    await bulletsNumbering.generateNumberingAsync(zip, this._numbering);
    await Style.generateStylesAsync(zip, this._style);

    zip.file("word/document.xml", xml);

    return await zip.generateAsync({ type: type });
  };
}

module.exports = DocxMerger;
