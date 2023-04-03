#pragma once

#include <QRunnable>
#include <QObject>
#include <QFile>
#include <QByteArray>

#include <opencv2/imgcodecs.hpp>
#include <opencv2/core/core.hpp>


class Add_Image_Thread : public QRunnable
{
    public:
        Add_Image_Thread();
        Add_Image_Thread(QString des_path, cv::Mat image);
        Add_Image_Thread(QString des_path, QString image_path);
        ~Add_Image_Thread();
        
    private:
        void run();

    private:
        cv::Mat image;
        QString des_path, image_path;
        bool is_mat;
};