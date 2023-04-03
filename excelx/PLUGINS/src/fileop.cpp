/***************************************************
File name:  fileop.cpp
Author: Li Xuan
Version: 1.0
Date:  2023-2-19
Description:  A function implementation of fileop.h
****************************************************/

#include "../include/fileop.h"

/*---------------
Public Functions
---------------*/

FileOp::FileOp()
{
   std::cout << "A FileOp instance has been created" << std::endl;
}

FileOp::~FileOp()
{
   std::cout << "A FileOp instance has been destroyed" << std::endl;
}

void FileOp::dirIterate(QString dir_path, void(*operation)(QString path, void *arg0, int flag), void *arg, int flag)
{
   QDirIterator it(dir_path);
   while(it.hasNext())
   {
      it.next();
      if(it.fileName() == "." || it.fileName() == "..") continue;
      if(it.fileInfo().isDir())
      {
         dirIterate(it.filePath(), operation, arg, (flag + 1));
      }
      (operation)(it.filePath(), arg, flag);
   }
}

void FileOp::zipFolder(const QString unziped_path)
{
   if(!QFile::exists(unziped_path))
   {
      qDebug() << unziped_path + " does't exists";
      return;
   }
   QString zip_path = unziped_path.mid(0, unziped_path.lastIndexOf('.')) + SUFFIX;
   if(QFile::exists(zip_path)) QFile::remove(zip_path);
   zipper::Zipper *zipper = new zipper::Zipper(zip_path.toStdString(), zipper::Zipper::openFlags::Overwrite);
   dirIterate(unziped_path, dirZiper, (void *) zipper, 0);
   zipper->close();
   delete zipper;
   zipper = nullptr;
   deleteCache(unziped_path);
}

QString FileOp::unzipFolder(QString ziped_path)
{
   int dot_pos = ziped_path.lastIndexOf('.');
   QString dir_path = ziped_path.mid(0, dot_pos) + ".cache";
   zipper::Unzipper *unzipper = new zipper::Unzipper(ziped_path.toStdString());
   bool ret = unzipper->extract(dir_path.toStdString());
   if(ret == false) return "";
   else return dir_path;
}

QString FileOp::unzipFolder(QString ziped_path, QString dest_path)
{  
   int dot_pos = dest_path.lastIndexOf('.');
   QString dir_path = dest_path.mid(0, dot_pos) + ".cache";
   QDir dir(dir_path);
   zipper::Unzipper *unzipper = new zipper::Unzipper(ziped_path.toStdString());
   unzipper->extract(dir_path.toStdString());
   return dir_path;
}
void FileOp::deleteCache(QString cache_path)
{
    QDir tmp_dir(cache_path);
    tmp_dir.removeRecursively();
}

/*---------------
Private Functions
---------------*/

void FileOp::dirZiper(QString path, void *arg, int flag)
{
   if(flag != 0) return;
   zipper::Zipper *zipper = reinterpret_cast<zipper::Zipper *>(arg);
   zipper->add(path.toStdString());
}

