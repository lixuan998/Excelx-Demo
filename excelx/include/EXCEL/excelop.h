#ifndef __EXCELOP_H__
#define __EXCELOP_H__

#include <vector>
#include <map>
#include <QDir>
#include <QImage>
#include <QFile>
#include <QFileInfo>
#include <QDebug>
#include <QCoreApplication>
#include <QThreadPool>

#include <opencv2/imgcodecs.hpp>
#include <opencv2/core/core.hpp>

#include "defs.h"

#include "xmlop.h"

#include "fileop.h"
#include "info.h"
#include "add_image_thread.h"

class ExcelOp
{
    //public functions
    public:
        ExcelOp();
        ExcelOp(QString src_model_path, QString dst_export_path = "");
        virtual ~ExcelOp();

        /*
            @brief 构造函数
            @param[in] src_model_path 模版文件地址
            @param[in] dst_export_path 导出文件地址
        */

        /*
            @brief 打开xlsx文件
            @param[in] src_model_path 模版文件地址
            @param[in] dst_export_path 导出文件地址
        */
        bool open(QString src_model_path = "", QString dst_export_path = "");

        /*
            @brief 关闭xlsx文件
        */
        void close();

        /*
            @breif 本函数的作用是在excel文件中添加一张表
            @param[in] sheet_model_sn 本参数的作用是决定添加的表的内容是否复制于另一张表。默认为空表，-1为上一张表的内容，可指定表的序号
        */
        bool AddSheet(int sheet_model_sn = 0);
        
        /*
            @breif 本函数的作用是在excel表中写入一个单元格
            @param[in] sheet_sn 写入的表的序号
            @param[in] cell_sn 写入的单元格序号
            @param[in] text 写入的内容
            @param[in] cell_style 写入的内容的样式
            @param[in] cell_height 写入的内容的高度
        */
        bool WriteCell(int sheet_sn, QString cell_sn, QString text, QString cell_style = "0", QString cell_height = "12.8");

        /*
            @breif 本函数的作用是在excel表中插入图片
            @param[in] sheet_sn 写入的表的序号
            @param[in] image_path 写入的图片地址
            @param[in] from_col 图片的起始列
            @param[in] from_row 图片的起始行
            @param[in] to_col 图片的终止列
            @param[in] to_row 图片的终止行
            @param[in] times 插入后的图片与原图的缩放比例
        */
        bool DrawCell(int sheet_sn, QString image_path, QString from_col, QString from_row, QString to_col, QString to_row, double times = DEFAULT_TIME);

        /*
            @breif 本函数的作用是替换字符串
            @param[in] mark 被替换的字符串
            @param[in] text 替换的字符串
        */
        int ReplaceSharedStringText(QString mark, QString text);

        /*
            @breif 本函数的作用是批量插入文字
            @param[in] sheet_sn 写入的表的序号
            @param[in] info 写入的信息结构体
            @param[in] direction 插入的方向
        */
        bool WriteBatch(int sheet_sn, Info info, int direction = VERTICAL);

        /*
            @breif 本函数的作用是批量插入图片
            @param[in] sheet_sn 写入的表的序号
            @param[in] info 写入的信息结构体
            @param[in] times 插入后的图片与原图的缩放比例
            @param[in] direction 插入的方向
        */
        bool DrawBatch(int sheet_sn, Info info, double times = DEFAULT_TIME, int direction = VERTICAL);

        void AddInfo(Info &info, QString label, QString text);
        void AddInfo(Info &info, QString label, cv::Mat mat);

        /*
            @breif 本函数的作用是获取表单数
        */
        int CountSheets();

        /*
            @breif 本函数的作用是在获取单元格的属性
            @param[in] sheet_sn 目标单元格所在表的序号
            @param[in] cell_sn 目标单元格标号
            @param[in] attribute_label 属性符号
        */
        QString GetCellAttribute(int sheet_sn, QString cell_sn, QString attribute_label);

        /*
            @breif 本函数的作用是在获取单元格的高度
            @param[in] sheet_sn 目标单元格所在表的序号
            @param[in] cell_sn 目标单元格标号
        */
        QString GetCellHeight(int sheet_sn, QString cell_sn);
    
    //private functions
    private:
        void Get_Image_SN();
        void Get_Sheet_SN();
        void Add_SharedString_Rels();
        int AddImage(int sheet_sn, QString image_path, int &image_height, int &image_width);

        int Get_String_SN(QString str);

        QString Get_Value_Type(QString value);

        std::vector<QString> CountRowAndCol(QString cell_sn);

    //public variables
    public:

    //private variables
    private:
        QString src_model_path;
        QString dst_export_path;
        QString cache_path;
        QString sheet_model_path;
        QString sheet_rels_model_path;
        QString workbook_xml_rels_text;
        QString workbook_text;
        QString content_type_text;

        QDir dir;
        int image_sn;
        int sheet_sn;

        int ret;
        
};
#endif
