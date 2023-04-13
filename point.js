// Scripting.FileSystemObject というオブジェクトを作成（JavaScript内でWSHを使ってファイルを扱う）
      var fs = new ActiveXObject("Scripting.FileSystemObject");

      // text.txtという新規のファイルを作成
      var file = fs.CreateTextFile("point.txt");

      // text.txtファイルへ書き込み
      file.Write(point);

      // text.txtファイルを閉じる
      file.Close();
