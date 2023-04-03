/*********************************************************************************************************************************************
File name:  fileop.h
Author: Li Xuan
Version: 1.0
Date:  2023-2-19
Description:  This module is used to zip and unzip files with suffix ".docx", it depends on zlib, zipper, minizips and Qt
Function List: 
    zipFolder:  Zip a folder
    unzipFolder:  Unzip a folder
    dirIterator:  Iterate a folder recursively. It also provide a function pointer argument that performs specific actions during the iteration
    deleteCache:  Delete the cache folder
    dirZipper:  A specific function that uses as a argument passses to the dirIterator function

***********************************************************************************************************************************************/

#ifndef __FILEOP_H__
#define __FILEOP_H__

#include "../zipper/zipper/zipper.h"
#include "../zipper/zipper/unzipper.h"
#include <QDebug>
#include <QString>
#include <string.h>
#include <string>
#include <assert.h>
#include <iostream>
#include <stdio.h>
#include <QFileInfo>
#include <QDirIterator>
#include <QDir>
#include <QFile>

#define MAX_PATH 1000
#define SUFFIX ".xlsx"


class FileOp{
    public:
        FileOp();
        ~FileOp();
        void static zipFolder(const QString ziped_path);

        QString static unzipFolder(QString ziped_path);
        
        QString static unzipFolder(QString ziped_path, QString dest_path);

        void static dirIterate(QString dir_path, void(*operation)(QString, void *, int ), void *arg, int flag);

        void static deleteCache(QString cache_path);

    private:

        void static dirZiper(QString path, void *arg, int flag);
    private:

};

#endif
