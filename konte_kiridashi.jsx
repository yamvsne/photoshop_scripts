/*

** 注意！！**
- パス名に日本語があるとうまく動作しないかも

*/

KETASUU = 3;

// 途中で抜けたいときのためにラベルをつける
main:
{
    // ファイル名取得
    var fileBaseName = activeDocument.name.split(".")[0];
    var saveDirPath = activeDocument.path + "/" + fileBaseName;

    try
    {
        // 選択範囲の情報取得
        var rectangle = activeDocument.selection.bounds;
    }
    catch (e)
    {
        // 選択範囲がない場合は何もせず終了
        alert("選択範囲を設定してね。");
        break main;
    }

    var w = rectangle[2] - rectangle[0];
    var h = rectangle[3] - rectangle[1];

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
    saveNum = getFileNum(fileList);
    var num = ('0000' + saveNum).slice(-KETASUU);
    saveFileName = saveDirPath + "/" + fileBaseName + "_" + num + ".jpg";
    file = new File(saveFileName);

    // ファイルを保存して新しく開いたウィンドウを閉じる
    option = setJpegOptions();
    newFile.saveAs(file, option, true, Extension.LOWERCASE);
    activeDocument.close(SaveOptions.DONOTSAVECHANGES);
}


function setJpegOptions()
{
    jpegOpt = new JPEGSaveOptions();
    jpegOpt.embedColorProfile = true;
    jpegOpt.quality = 12;

    return jpegOpt;
}

function getFileNum(fileList)
{
    // \D以降があることで、最後にマッチしたものに限定できる
    re = new RegExp("(\\d{" + KETASUU + "})\\.\\D{" + KETASUU + "}");

    var cnt;
    var max = -1; // ファイルが存在しないときは0が返る

    for (cnt=0; cnt<fileList.length; cnt++)
    {
        filename = fileList[cnt].name;
        num = filename.match(re);

        if (num != null)
        {
            num = Number(num[1]);
            if (num > max)
            {
                max = num;
            }
        }
    }
    // 存在するファイル名の連番で最も大きい数+1を返す
    return max + 1;
}
