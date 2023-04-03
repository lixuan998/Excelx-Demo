#pragma once
#include <QString>
#include <string>
#include <vector>
#include <map>
#include <QFile>
#include <opencv2/core/core.hpp>
#include <iostream>

class Info
{
    public:
        Info();
        ~Info();
        void addInfo(QString label, QString text);
        void addInfo(QString label, cv::Mat image);
        void clear();
        std::map<QString, std::vector<QString>> getLabelText();
        std::map<QString, std::vector<QString>> getLabelImagePath();
        std::map<QString, std::vector<cv::Mat>> getLabelImage();
    private:
        std::map<QString, std::vector<QString>> label_to_text_map, label_to_image_path_map;
        std::map<QString, std::vector<cv::Mat>> label_to_mat_map;
};