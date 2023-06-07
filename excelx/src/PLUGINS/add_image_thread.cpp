#include "add_image_thread.h"

Add_Image_Thread::Add_Image_Thread()
{
    
}
Add_Image_Thread::Add_Image_Thread(QString des_path, cv::Mat image)
{
    this->setAutoDelete(true);
    this->des_path = des_path;
    this->image = image;
    is_mat = true;
}

Add_Image_Thread::Add_Image_Thread(QString des_path, QString image_path)
{
    this->setAutoDelete(true);
    this->des_path = des_path;
    this->image_path = image_path;
    is_mat = false;
}

Add_Image_Thread::~Add_Image_Thread()
{
    
}

void Add_Image_Thread::run()
{
    if(is_mat == true)
    {
        cv::imwrite(des_path.toStdString(), image);
    }
    else
    {
        QFile f_mark, f_image;
        f_mark.setFileName(des_path);
	    f_image.setFileName(image_path);
        f_mark.open(QIODevice::ReadWrite);
	    f_image.open(QIODevice::ReadWrite);
    
	    QByteArray b_array = f_image.readAll();
	    f_mark.write(b_array);

        f_mark.close();
        f_image.close();
    }
}