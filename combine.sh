#!/bin/sh

git diff --exit-code --quiet ./bin
if [ $? -eq 1 ]; then
  echo 'Cannot combine! (Diff exists in VBA. You may overwrite pre-decombined VBA!)'
  return 0
fi

# ユーザ名の取得 -> userFolderName
userFolderPath="`cmd.exe /c echo %USERPROFILE% | sed 's/\r//g'`"
userFolderName=${userFolderPath##*\\}

# リポジトリ名の取得 -> repoName
repoPath="`pwd`"
repoName=${repoPath##*/}

# Windows Script Fileが実行できる環境先
copyToPath="/mnt/c/Users/${userFolderName}/Downloads/"
# リポジトリの複製
cp -rf ../$repoName $copyToPath

# 複製先に移動
cd $copyToPath

# Windows Script Fileの実行
cmd.exe /c cscript ${repoName}/vbac.wsf combine

# 生成されたファイルをリポジトリに複製
cp -f ${repoName}/bin/*.xlsm $repoPath/bin
