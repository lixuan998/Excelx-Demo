#pragma once
#include <QString>
#include <QFile>
#include <QTextStream>
#include <QDir>
#include <QDebug>
#include <QCoreApplication>

#include "../../PLUGINS/include/info.h"
#include "../../EXCEL/include/defs.h"

class XmlOp{
    //public functions
    public:
        XmlOp();
        ~XmlOp();

        static bool LoadXml(const QString fileName, QString &output);
        static bool SaveXml(const QString fileName, const QString &input);
        static int AnalyzeXmlLabels(const QString &input, std::vector<QString> &output, QString label);
        static int CountLabels(const QString xml_text, QString label);
        static bool ReplaceText(QString &xml_text, const QString &mark, const QString &replace);
        static int AddRelationship(QString &xml_text, QString new_relationship);
        static bool AddWorkBookSheet(QString &xml_text, int sheet_sn, int rid_sn);
        static bool AddContentType(QString &xml_text, QString new_content_type);
        static bool AddText(QString &xml_text, QString new_text, QString label, int label_index = 1);
        
        static bool ModifyAttributes(QString &xml_text, QString xml_label, QString xml_attribute_label, QString xml_attribute_value, int xml_label_index = 1);
        static QString GetAttribute(QString &xml_text, QString xml_label, QString xml_attribute_label, QString xml_feature_attribute_label);
        static QString GetCellHeight(QString &xml_text, QString cell_sn);
        static int CountStringSn(QString &xml_text, QString str);

        static int Find(QString src, QString des, int index = 1);
        static void ExtractTexts(std::vector<QString> &src);

        static bool ReplaceSharedStringText(QString &xml_text, QString label, QString text);
        static std::map<QString, std::vector<QString>> GetCellSns(QString &xml_text, Info info, int direction);

    //private functions
    private:
        static void RemoveSpaces(QString &input);
        static void RemoveOuterQuotation(QString &input);
        static bool MergeTexts(QString &xml_text, QString &label);
        static QString GetCellSn(QString &xml_text, QString cell_text);
        static QString NextCellSn(QString current_cell_sn, int direction = VERTICAL);

};