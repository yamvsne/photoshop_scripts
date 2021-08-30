/*

** 注意！！**
- パス名に日本語があるとうまく動作しないかも
- ファイルの数で連番を振っているので、ファイルを削除したりして数が変化するとうまく連番が振れないことがあります（いずれ直したい）

*/

// ファイル名取得
var fileBaseName = activeDocument.name.split(".")[0];
var saveDirPath = activeDocument.path + "/" + fileBaseName;

// 選択範囲の情報取得
var rectangle = activeDocument.selection.bounds;
var w = rectangle[2] - rectangle[0];
var h = rectangle[3] - rectangle[1];
// alert("w: " + w + "h: " + h);

// 選択範囲をコピーして新しいドキュメントに貼り付け
activeDocument.selection.copy();
newFile = app.documents.add(w,h);
newFile.paste();

// 保存先フォルダ作成
saveFolder = new Folder(saveDirPath);
if (saveFolder.exists == false)
{
    saveFolder.create();    
}

// 連番にするためにフォルダにあるファイル数を取得(3桁0埋め)
fileList = saveFolder.getFiles();
var num = ('0000' + fileList.length).slice(-3);
saveFileName = saveDirPath + "/" + fileBaseName + "_" + num + ".jpg"; 
file = new File(saveFileName);

// もしファイルが存在していたらアラートを出す
if (file.exists == true)
{
    alert("file: " + saveFileName + "is exist! Overwrite.");
}

// ファイルを保存して新しく開いたウィンドウを閉じる
option = setJpegOptions();
newFile.saveAs(file, option, true, Extension.LOWERCASE);
activeDocument.close(SaveOptions.DONOTSAVECHANGES);


// 
function setJpegOptions()
{
    jpegOpt = new JPEGSaveOptions();
    jpegOpt.embedColorProfile = true;
    jpegOpt.quality = 12;

    return jpegOpt;
}
