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

#include "../include/defs.h"

#include "../XML/include/xmlop.h"

#include "../PLUGINS/include/fileop.h"
#include "../PLUGINS/include/info.h"
#include "../PLUGINS/include/add_image_thread.h"

class ExcelOp
{
    //public functions
    public:
        ExcelOp();
        ExcelOp(QString src_model_path, QString dst_export_path = "");
        ~ExcelOp();

        bool open(QString src_model_path = "", QString dst_export_path = "");
        void close();

        int CountSheets();

        bool AddSheet(int sheet_model_sn = 0);

        void AddInfo(Info &info, QString label, QString text);
        void AddInfo(Info &info, QString label, cv::Mat mat);

        bool WriteCell(int sheet_sn, QString cell_sn, QString text, QString cell_style = "0", QString cell_height = "12.8");
        bool DrawCell(int sheet_sn, QString from_col, QString from_row, QString to_col, QString to_row, double times = 1);

        QString GetCellAttribute(int sheet_sn, QString cell_sn, QString attribute_label);

        QString GetCellHeight(int sheet_sn, QString cell_sn);

        int ReplaceSharedStringText(QString mark, QString text);

        bool WriteBatch(int sheet_sn, Info info, int direction = VERTICAL);
    
    //private functions
    private:
        void Get_Image_SN();
        void Get_Sheet_SN();
        void Add_SharedString_Rels();
        

        int Get_String_SN(QString str);

        QString Get_Value_Type(QString value);

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
