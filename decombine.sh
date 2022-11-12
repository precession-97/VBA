#!/bin/sh

# ユーザ名の取得 -> userFolderName
userFolderPath="`cmd.exe /c echo %USERPROFILE% | sed 's/\r//g'`"
userFolderName=${userFolderPath##*\\}

# リポジトリ名の取得 -> repoName
repoPath="`pwd`"
repoName=${repoPath##*/}

# Windows Script Fileが実行できる環境先
copyToPath="/mnt/c/Users/${userFolderName}/Downloads/"
# リポジトリの複製
cp -rf $repoPath $copyToPath

# 複製先に移動
cd $copyToPath

# Windows Script Fileの実行
cmd.exe /c cscript ${repoName}/vbac.wsf decombine

# 生成されたファイルをリポジトリに複製
cp -rf ${repoName}/src $repoPath
