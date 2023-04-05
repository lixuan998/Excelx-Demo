#include "../include/excelop.h"

/*---------------
Public Functions
---------------*/

ExcelOp::ExcelOp()
{
    qDebug() << "ExcelOp instance created";
}

ExcelOp::ExcelOp(QString src_model_path, QString dst_export_path)
{
    this->src_model_path = src_model_path;
    if(dst_export_path.isEmpty()) this->dst_export_path = src_model_path;
    else this->dst_export_path = dst_export_path;

    qDebug() << "ExcelOp instance created";
}

ExcelOp::~ExcelOp()
{
     qDebug() << "ExcelOp instance distroyed";
}

bool ExcelOp::open(QString src_model_path, QString dst_export_path)
{
    if(!src_model_path.isEmpty())
    {
        this->src_model_path = src_model_path;
        if(!dst_export_path.isEmpty()) this->dst_export_path = dst_export_path;
        else this->dst_export_path = src_model_path;
    }
    
    if(this->src_model_path.isEmpty() || this->dst_export_path.isEmpty())
    {
        qDebug() << "Error: src_model_path or dst_export_path is empty";
        return false;
    }

    if(!QFile::exists(this->src_model_path))
    {
        qDebug() << "Error: " + this->src_model_path + " does not exist";
        return false;
    }

    cache_path = FileOp::unzipFolder(this->src_model_path, this->dst_export_path);

    if(cache_path.isEmpty())
    {
        qDebug() << "Error: cache_path is empty, fail to unzip " + this->src_model_path;
        return false;
    }

    Get_Sheet_SN();
    return true;
}

void ExcelOp::close()
{
    FileOp::zipFolder(cache_path);
    cache_path.clear();
}

int ExcelOp::CountSheets()
{
    return sheet_sn;
}

bool ExcelOp::AddSheet(int sheet_model_sn)
{
    QString sheet_rels_dst_path;
    if(sheet_model_sn == 0) sheet_model_path = XML_MODEL_PATH + "default_sheet.xml";
    else if(sheet_model_sn == -1)
    {
        sheet_model_path = cache_path + WORKSHEETS_PATH + "sheet" + QString::number(sheet_sn) + ".xml";
        sheet_rels_model_path = cache_path + WORKSHEETS_RELS_PATH + "sheet" + QString::number(sheet_sn) + ".xml.rels";
        if(QFile::exists(sheet_model_path)) sheet_rels_dst_path = cache_path + WORKSHEETS_RELS_PATH + "sheet" + QString::number(sheet_sn + 1) + ".xml.rels";
    }
    else
    {
        sheet_model_path = cache_path + WORKSHEETS_PATH + "sheet" + QString::number(sheet_model_sn) + ".xml";
        if(QFile::exists(sheet_model_path)) sheet_rels_dst_path = cache_path + WORKSHEETS_RELS_PATH + "sheet" + QString::number(sheet_sn + 1) + ".xml.rels";
    }
    QString sheet_text, sheet_rels_text;
    ret = XmlOp::LoadXml(sheet_model_path, sheet_text);
    if(ret == false)
    {
        qDebug() << "fail to load file: " << sheet_model_path;
        return false;
    }
    QString sheet_des_path = cache_path + WORKSHEETS_PATH + "sheet" + QString::number(++ sheet_sn) + ".xml";
    ret = XmlOp::SaveXml(sheet_des_path, sheet_text);
    if(ret == false)
    {
        qDebug() << "fail to save file: " << sheet_des_path;
        return false;
    }
    if(!sheet_rels_dst_path.isEmpty())
    {
        ret = XmlOp::LoadXml(sheet_rels_model_path, sheet_rels_text);
        if(ret == false)
        {
            qDebug() << "fail to load file: " << sheet_rels_model_path;
            return false;
        }
        ret = XmlOp::SaveXml(sheet_rels_dst_path, sheet_rels_text);
        if(ret == false)
        {
            qDebug() << "fail to save file: " << sheet_rels_dst_path;
            return false;
        }
    }
    QString workbook_xml_rels_path = cache_path + WORKBOOK_RELS_PATH + "workbook.xml.rels";
    ret = XmlOp::LoadXml(workbook_xml_rels_path, workbook_xml_rels_text);
    if(ret == false)
    {
        qDebug() << "fail to load file: " << workbook_xml_rels_path;
        return false;
    }
    int count = XmlOp::CountLabels(workbook_xml_rels_text, "Relationship");
    
    QString new_relationship = "<Relationship Id=\"rId${RID_SN}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet${SHEET_SN}.xml\" />";

    int ret_rid_sn = XmlOp::CountLabels(workbook_xml_rels_text, "Relationship");
    ret = XmlOp::ReplaceText(new_relationship, "${RID_SN}", QString::number(++ ret_rid_sn));
    if(ret == false)
    {
        qDebug() << "Error: fail to replace RID_SN";
        return -1;
    }
    ret = XmlOp::ReplaceText(new_relationship, "${SHEET_SN}", QString::number(sheet_sn));
    if(ret == false)
    {
        qDebug() << "Error: fail to replace SHEET_SN";
        return -1;
    }
    ret = XmlOp::AddRelationship(workbook_xml_rels_text, new_relationship);
    if(ret == -1)
    {
        qDebug() << "AddRels fail";
        return false;
    }
    ret = XmlOp::SaveXml(workbook_xml_rels_path, workbook_xml_rels_text);
    if(ret == false)
    {
        qDebug() << "fail to save file: " << workbook_xml_rels_path;
        return false;
    }

    QString workbook_path = cache_path + XL_PATH + "workbook.xml";
    ret = XmlOp::LoadXml(workbook_path, workbook_text);
    if(ret == false)
    {
        qDebug() << "fail to load file: " << workbook_path;
        return false;
    }

    XmlOp::AddWorkBookSheet(workbook_text, sheet_sn, ret_rid_sn);
    ret = XmlOp::SaveXml(workbook_path, workbook_text);
    if(ret == false)
    {
        qDebug() << "fail to save file: " << workbook_path;
        return false;
    }
    QString content_type_path = cache_path + "/[Content_Types].xml";
    ret = XmlOp::LoadXml(content_type_path, content_type_text);
    if(ret == false)
    {
        qDebug() << "fail to load file: " << content_type_path;
        return false;
    }
    QString new_content_type = "<Override PartName=\"/xl/worksheets/sheet${SHEET_SN}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\" />";

    ret = XmlOp::ReplaceText(new_content_type, "${SHEET_SN}", QString::number(sheet_sn));
    if(ret == false)
    {
        qDebug() << "Error: fail to replace SHEET_SN";
        return -1;
    }
    ret = XmlOp::AddContentType(content_type_text, new_content_type);
    if(ret == false)
    {
        qDebug() << "AddContentType fail";
        return false;
    }
    ret = XmlOp::SaveXml(content_type_path, content_type_text);
    if(ret == false)
    {
        qDebug() << "fail to save file: " << content_type_path;
        return false;
    }

    return true;
}

void ExcelOp::AddInfo(Info &info, QString label, QString text)
{
    bool isok = true;
    label.toInt(&isok);
    if(isok == false)
    label = QString::number(Get_String_SN(label));
    info.addInfo(label, text);
}

void ExcelOp::AddInfo(Info &info, QString label, cv::Mat mat)
{
    info.addInfo(label, mat);
}

bool ExcelOp::WriteCell(int sheet_sn, QString cell_sn, QString text, QString cell_style, QString cell_height)
{
    int string_sn = Get_String_SN(text);

    QString sheet_text, sheet_file_path, new_cell_text;
    sheet_file_path = cache_path + WORKSHEETS_PATH + "sheet" + QString::number(sheet_sn) + ".xml";
    ret = XmlOp::LoadXml(XML_MODEL_PATH + "new_cell.xml", new_cell_text);
    if(ret == false)
    {
        qDebug() << "fail to load " + XML_MODEL_PATH + "new_cell.xml";
        return false;
    }
    ret = XmlOp::LoadXml(sheet_file_path, sheet_text);
    int tmp_pos = XmlOp::Find(sheet_text, "<sheetData/>");
    if(tmp_pos != -1) sheet_text.replace("<sheetData/>", "<sheetData></sheetData>");
    if(ret == false)
    {
        qDebug() << "fail to load " + sheet_file_path;
        return false;
    }
    QString value_type = Get_Value_Type(text);

    new_cell_text.replace("${TYPE}", value_type);

    if(value_type == "s") new_cell_text.replace("${VALUE}", QString::number(string_sn));
    else if(value_type == "n") new_cell_text.replace("${VALUE}", text);

    new_cell_text.replace("${CELL}", cell_sn);

    for(int i = 0; i < cell_sn.size(); ++ i)
    {
        if(cell_sn[i].isNumber())
        {
            new_cell_text.replace("${ROW}", cell_sn.mid(i));
            break;
        }
    }

    new_cell_text.replace("${STYLE}", cell_style);
    new_cell_text.replace("${HEIGHT}", cell_height);

    int insert_pos = XmlOp::Find(sheet_text, "</sheetData>");
    sheet_text.insert(insert_pos, new_cell_text);

    XmlOp::SaveXml(sheet_file_path, sheet_text);
    return false;
}

bool ExcelOp::DrawCell(int sheet_sn, QString image_path, QString from_col, QString from_row, QString to_col, QString to_row, double times)
{
    ret = 1;
    QString sheet_path = cache_path + WORKSHEETS_PATH + "sheet" + QString::number(sheet_sn) + ".xml";
    if(!QFile::exists(sheet_path))
    {
        qDebug() << "Sheet " << sheet_sn << " does not exist";
        return false;
    }
    QString xml_rels_path = cache_path + WORKSHEETS_RELS_PATH;
    QString drawings_path = cache_path + XL_PATH + "drawings/";
    QString drawing_rels_path = drawings_path + "_rels/";
    QString media_path = cache_path + MEDIA_PATH;
    if(!QFile::exists(drawings_path)) dir.mkpath(drawings_path);
    if(!QFile::exists(drawing_rels_path)) dir.mkpath(drawing_rels_path);
    if(!QFile::exists(media_path)) dir.mkpath(media_path);
    if(!QFile::exists(xml_rels_path)) dir.mkpath(xml_rels_path);

    QString sheetn_xml_rels_path = xml_rels_path + "sheet" + QString::number(sheet_sn) + ".xml.rels";
    QString drawingn_xml_rels_path = drawing_rels_path + "drawing" + QString::number(sheet_sn) + ".xml.rels";
    QString drawingn_path = drawings_path + "drawing" + QString::number(sheet_sn) + ".xml";
    QString content_types_text;
    ret &= XmlOp::LoadXml(cache_path + "/[Content_Types].xml", content_types_text);
    if(!QFile::exists(sheetn_xml_rels_path))
    {
        QString tmp;
        QString sheetn_xml_rels_model_path = XML_MODEL_PATH + "new_xml_rels.xml";
        ret &= XmlOp::LoadXml(sheetn_xml_rels_model_path, tmp);
        ret &= XmlOp::SaveXml(sheetn_xml_rels_path, tmp);
        QString new_content_type = "<Override PartName=\"/xl/worksheets/_rels/sheet" + QString::number(sheet_sn) + ".xml.rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" />";
        ret &= XmlOp::AddContentType(content_types_text, new_content_type);
    }
    if(!QFile::exists(drawingn_xml_rels_path))
    {
        QString tmp;
        QString drawingn_xml_rels_model_path = XML_MODEL_PATH + "new_xml_rels.xml";
        ret &= XmlOp::LoadXml(drawingn_xml_rels_model_path, tmp);
        ret &= XmlOp::SaveXml(drawingn_xml_rels_path, tmp);
        QString new_content_type = "<Override PartName=\"/xl/drawings/_rels/drawing" + QString::number(sheet_sn) + ".xml.rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" />";
        ret &= XmlOp::AddContentType(content_types_text, new_content_type);
    }
    if(!QFile::exists(drawingn_path))
    {
        
        QString tmp;
        QString drawingn_model_path = XML_MODEL_PATH + "new_drawing.xml";
        ret &= XmlOp::LoadXml(drawingn_model_path, tmp);
        ret &= XmlOp::SaveXml(drawingn_path, tmp);
        QString new_content_type = "<Override PartName=\"/xl/drawings/drawing" + QString::number(sheet_sn) + ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\" />";
        ret &= XmlOp::AddContentType(content_types_text, new_content_type);
    }
    ret &= XmlOp::SaveXml(cache_path + "/[Content_Types].xml", content_types_text);
    if(!QFile::exists(image_path))
    {
        qDebug() << "Image path " + image_path + " does not exist";
        return false;
    }
    Get_Image_SN();
    int image_height, image_width;
    int image_rid =  AddImage(sheet_sn, image_path, image_height, image_width);
    if(!image_rid) return false;
    image_height *= times;
    image_width *= times;
    QString new_drawing_item_text, drawingn_text;
    
    ret &= XmlOp::LoadXml(XML_MODEL_PATH + "new_drawing_item.xml", new_drawing_item_text);
    ret &= XmlOp::LoadXml(drawingn_path, drawingn_text);

    new_drawing_item_text.replace("${IMAGE_COL_FROM}", from_col);
    new_drawing_item_text.replace("${IMAGE_ROW_FROM}", from_row);
    new_drawing_item_text.replace("${IMAGE_COL_TO}", to_col);
    new_drawing_item_text.replace("${IMAGE_ROW_TO}", to_row);
    new_drawing_item_text.replace("${IMAGE_COL_OFF}", QString::number(image_width));
    new_drawing_item_text.replace("${IMAGE_ROW_OFF}", QString::number(image_height));
    new_drawing_item_text.replace("${IMAGE_ID}", QString::number(image_rid));
    new_drawing_item_text.replace("${IMAGE_RID}", "rId" + QString::number(image_rid));
    
    ret &= XmlOp::AddText(drawingn_text, new_drawing_item_text, "</xdr:wsDr");
    ret &= XmlOp::SaveXml(drawingn_path, drawingn_text);

    QString sheetn_xml_rels_text;
    XmlOp::LoadXml(sheetn_xml_rels_path, sheetn_xml_rels_text);
    int ret_rid_sn = XmlOp::CountLabels(sheetn_xml_rels_text, "Relationship");

    QString new_sheetn_xml_rels = "<Relationship Id=\"rId" + QString::number(++ ret_rid_sn) + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing" + QString::number(sheet_sn) + ".xml\"/>";
    ret &= XmlOp::AddRelationship(sheetn_xml_rels_text, new_sheetn_xml_rels);

    ret &= XmlOp::SaveXml(sheetn_xml_rels_path, sheetn_xml_rels_text);

    QString sheetn_addon = "<drawing r:id=\"rId" + QString::number(ret_rid_sn) + "\"/>";
    QString sheetn_text;

    ret &= XmlOp::LoadXml(sheet_path, sheetn_text);
    ret &= XmlOp::AddText(sheetn_text, sheetn_addon, "</worksheet>");
    ret &= XmlOp::SaveXml(sheet_path, sheetn_text);
    return ret;
}

QString ExcelOp::GetCellAttribute(int sheet_sn, QString cell_sn, QString attribute_label)
{
    QString sheet_text, sheet_file_path;
    sheet_file_path = cache_path + WORKSHEETS_PATH + "sheet" + QString::number(sheet_sn) + ".xml";
    ret = XmlOp::LoadXml(sheet_file_path, sheet_text);
    if(ret == false)
    {
        qDebug() << "fail to load " + sheet_file_path;
        return "";
    }
    QString str_attribute = XmlOp::GetAttribute(sheet_text, "c", attribute_label, cell_sn);
    if(str_attribute == "")
    {
        qDebug() << "fail to get cell attribute " + attribute_label;
        return "";
    }
    XmlOp::SaveXml(sheet_file_path, sheet_text);
    return str_attribute;
}

QString ExcelOp::GetCellHeight(int sheet_sn, QString cell_sn)
{
    QString sheet_text, sheet_file_path;
    sheet_file_path = cache_path + WORKSHEETS_PATH + "sheet" + QString::number(sheet_sn) + ".xml";
    ret = XmlOp::LoadXml(sheet_file_path, sheet_text);
    if(ret == false)
    {
        qDebug() << "fail to load " + sheet_file_path;
        return "";
    }
    QString height = XmlOp::GetCellHeight(sheet_text, cell_sn);
    XmlOp::SaveXml(sheet_file_path, sheet_text);
    return height;
}

int ExcelOp::ReplaceSharedStringText(QString mark, QString text)
{
    QString str_file_path = (cache_path + XL_PATH + "sharedStrings.xml");
    QString sharedstrings_text;
    if(!QFile::exists(str_file_path))
    {
        qDebug() << "sharedStrings.xml does not exist, can not replace text";
        return false;
    }
    XmlOp::LoadXml(str_file_path, sharedstrings_text);
    mark = XmlOp::GetSharedString(sharedstrings_text, mark);
    qDebug() << "text: " << text;
    ret = XmlOp::ReplaceSharedStringText(sharedstrings_text, mark, text);
    ret = XmlOp::SaveXml(str_file_path, sharedstrings_text);
    return ret;
}

bool ExcelOp::WriteBatch(int sheet_sn, Info info, int direction)
{
    QString sheet_text, sheet_file_path;
    sheet_file_path = cache_path + WORKSHEETS_PATH + "sheet" + QString::number(sheet_sn) + ".xml";
    ret = XmlOp::LoadXml(sheet_file_path, sheet_text);
    int tmp_pos = XmlOp::Find(sheet_text, "<sheetData/>");
    if(tmp_pos != -1) sheet_text.replace("<sheetData/>", "<sheetData></sheetData>");
    if(ret == false)
    {
        qDebug() << "fail to load " + sheet_file_path;
        return false;
    }
    std::map<QString, std::vector<QString>> label_to_cell_sn_map = XmlOp::GetCellSns(sheet_text, info, direction);
    ret = XmlOp::SaveXml(sheet_file_path, sheet_text);
    std::map<QString, std::vector<QString>> label_to_text_map = info.getLabelText();
    for(auto it = label_to_text_map.begin(); it != label_to_text_map.end(); ++ it)
    {
        QString tmp_label = it->first;
        QString style = GetCellAttribute(sheet_sn, label_to_cell_sn_map[tmp_label][0], "s");
        QString height = GetCellHeight(sheet_sn, label_to_cell_sn_map[tmp_label][0]);
        ReplaceSharedStringText(tmp_label, label_to_text_map[tmp_label][0]); 
        for(int i = 1; i < it->second.size(); ++ i)
        {
            WriteCell(sheet_sn, label_to_cell_sn_map[tmp_label][i], label_to_text_map[tmp_label][i], style, height);
        }
    }
    
    return true;
}

bool ExcelOp::DrawBatch(int sheet_sn, Info info, double times, int direction)
{
    QString sheet_text, sheet_file_path;
    sheet_file_path = cache_path + WORKSHEETS_PATH + "sheet" + QString::number(sheet_sn) + ".xml";
    ret = XmlOp::LoadXml(sheet_file_path, sheet_text);
    int tmp_pos = XmlOp::Find(sheet_text, "<sheetData/>");
    if(tmp_pos != -1) sheet_text.replace("<sheetData/>", "<sheetData></sheetData>");
    if(ret == false)
    {
        qDebug() << "fail to load " + sheet_file_path;
        return false;
    }
    std::map<QString, std::vector<QString>> label_to_cell_sn_map = XmlOp::GetCellSns(sheet_text, info, direction);
    ret = XmlOp::SaveXml(sheet_file_path, sheet_text);
    std::map<QString, std::vector<QString>> label_to_image_path_map = info.getLabelImagePath();

    for(auto it = label_to_image_path_map.begin(); it != label_to_image_path_map.end(); ++ it)
    {
        QString tmp_label = it->first;
        ReplaceSharedStringText(tmp_label, "");
        QString style = GetCellAttribute(sheet_sn, label_to_cell_sn_map[tmp_label][0], "s");
        QString height = GetCellHeight(sheet_sn, label_to_cell_sn_map[tmp_label][0]);
        for(int i = 0; i < it->second.size(); ++ i)
        {
            WriteCell(sheet_sn, label_to_cell_sn_map[tmp_label][i], "", style, height);
            std::vector<QString> col_and_row = CountRowAndCol(label_to_cell_sn_map[tmp_label][i]);
            DrawCell(sheet_sn, label_to_image_path_map[tmp_label][i], col_and_row[0], col_and_row[1], col_and_row[0], col_and_row[1], times);
        }
    }
    return true;
}

void ExcelOp::Get_Image_SN()
{
    QDir tempdir((cache_path + MEDIA_PATH));
    image_sn = tempdir.count() - 2;
}

void ExcelOp::Get_Sheet_SN()
{
    int cnt = 0;
    QDir tempdir((cache_path + WORKSHEETS_PATH));
    for(int i = 0; i < tempdir.count(); ++ i)
    {
        QFileInfo fileinfo(tempdir[i]);
        if(fileinfo.isDir()) continue;
        else if(fileinfo.suffix() == "xml") cnt ++;
    }
    sheet_sn = cnt;
}

void ExcelOp::Add_SharedString_Rels()
{
    std::vector<QString> rels_list;
    QString workbook_rels_file_path = cache_path + WORKBOOK_RELS_PATH + "workbook.xml.rels";
    QString workbook_rels_text;
    XmlOp::LoadXml(workbook_rels_file_path, workbook_rels_text);
    int rid = XmlOp::AnalyzeXmlLabels(workbook_rels_text, rels_list, "Relationship");
    QString new_sharedstring_rels_text;
    new_sharedstring_rels_text = "<Relationship Id=\"rId${RID}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>";
    new_sharedstring_rels_text.replace("${RID}", QString::number(rid));
    int insert_pos = XmlOp::Find(workbook_rels_text, "</Relationships>");
    workbook_rels_text.insert(insert_pos, new_sharedstring_rels_text);
    XmlOp::SaveXml(workbook_rels_file_path, workbook_rels_text);
}

int ExcelOp::AddImage(int sheet_sn, QString image_path, int &image_height, int &image_width)
{
    int ret = 1;
    cv::Mat mat = cv::imread(image_path.toStdString());
    QString media_path = cache_path + MEDIA_PATH;
    QString new_image_file_name = "image" + QString::number(++image_sn) + ".jpeg";
    ret &= cv::imwrite((media_path + new_image_file_name).toStdString(), mat);
    image_height = mat.rows;
    image_width = mat.cols;
    if(!ret)
    {
        qDebug() << "Fail to write image " << image_path;
        return 0;
    }
    QString tmp_text;
    ret &= XmlOp::LoadXml(cache_path + "/[Content_Types].xml", tmp_text);
    QString new_content_type = "<Override PartName=\"/xl/media/image" + QString::number(image_sn) + ".jpeg\" ContentType=\"image/jpeg\" />";
    ret &= XmlOp::AddContentType(tmp_text, new_content_type);
    ret &= XmlOp::SaveXml(cache_path + "/[Content_Types].xml", tmp_text);

    QString drawings_path = cache_path + XL_PATH + "drawings/";
    QString drawing_rels_path = drawings_path + "_rels/";
    QString drawingn_xml_rels_path = drawing_rels_path + "drawing" + QString::number(sheet_sn) + ".xml.rels";
    ret &= XmlOp::LoadXml(drawingn_xml_rels_path, tmp_text);

    int rid_num = XmlOp::CountLabels(tmp_text, "Relationship");

    QString new_relationship_type = "<Relationship Id=\"rId" + QString::number(++ rid_num) + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/" + new_image_file_name + "\" />";
    ret &= XmlOp::AddRelationship(tmp_text, new_relationship_type);
    ret &= XmlOp::SaveXml(drawingn_xml_rels_path, tmp_text);
    if(!ret) return ret;
    return rid_num;
}

int ExcelOp::Get_String_SN(QString str)
{
    QString str_file_path = (cache_path + XL_PATH + "sharedStrings.xml");
    QString sharedstrings_text;
    if(!QFile::exists(str_file_path))
    {
        QString new_shared_string_model, new_shared_string_model_path;
        new_shared_string_model_path = XML_MODEL_PATH + "new_sharedstring.xml";
        XmlOp::LoadXml(new_shared_string_model_path, new_shared_string_model);
        XmlOp::SaveXml(str_file_path, new_shared_string_model);
        Add_SharedString_Rels();
        sharedstrings_text = new_shared_string_model;
    }
    else
    {
        XmlOp::LoadXml(str_file_path, sharedstrings_text);
    }
    int counts = XmlOp::CountStringSn(sharedstrings_text, str);
    XmlOp::SaveXml(str_file_path, sharedstrings_text);
    return counts;
}

QString ExcelOp::Get_Value_Type(QString value)
{
    return "s";
    bool isok = true;
    value.toInt(&isok);
    if(isok) return "n";
    else return "s";
}

std::vector<QString> ExcelOp::CountRowAndCol(QString cell_sn)
{
    std::vector<QString> res(2);
    QString col, row;
    for(int i = 0; i < cell_sn.length(); ++ i)
    {
        if(cell_sn[i].isDigit())
        {
            col = cell_sn.mid(0, i);
            row = cell_sn.mid(i);
            break;
        }
    }
    row = QString::number(row.toInt() - 1);
    res[1] = row;
    int count_col = 0;
    QString begin_col = "A";
    QString alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

    while(begin_col != col)
    {
        int addon = 1;
        int cur_index = begin_col.size() - 1;
        while(addon)
        {
            if(cur_index  == -1)
            {
                begin_col = alphabet[-1 + addon] + begin_col;
                addon = 0;
                continue;
            }
            int cur_num;
            for(int i = 0; i < 26; ++ i)
            {
                if((begin_col[cur_index]) == alphabet[i])
                {
                    cur_num = i;
                    break;
                }
            }
            begin_col[cur_index] = alphabet[(cur_num + addon) % 26];
            addon = (cur_num + addon) / 26;
            -- cur_index;
        }
        ++ count_col;
    }
    

    col = QString::number(count_col);
    res[0] = col;

    return res;
    
}