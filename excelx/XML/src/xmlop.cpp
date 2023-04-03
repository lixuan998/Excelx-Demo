#include "../include/xmlop.h"

XmlOp::XmlOp()
{

}

XmlOp::~XmlOp()
{

}

bool XmlOp::LoadXml(const QString fileName, QString &output)
{
    QFile file(fileName);
    if (!file.open(QIODevice::ReadOnly | QIODevice::Text))
    {
        return false;
    }

    QTextStream stream(&file);
    output = stream.readAll();
    file.close();
    return true;
}

bool XmlOp::SaveXml(const QString fileName, const QString &input)
{
    QFile file(fileName);
    if (!file.open(QIODevice::WriteOnly | QIODevice::Text))
    {
        return false;
    }

    QTextStream stream(&file);
    stream << input;
    file.close();
    return true;
}

int XmlOp::AnalyzeXmlLabels(const QString &input, std::vector<QString> &output, QString label)
{
    output.clear();
    int pos_start, pos_end, index = 0;
    int tmp_pos_before = 0;
    int tmp_pos_after = 0;
    pos_start = -1;
    pos_start = Find(input, ("<" + label), index + 1);
    pos_end = Find(input, ("</" + label + ">"), index + 1);
    if(pos_end == -1) pos_end = Find(input, ("/>"), index + 1);
    while(pos_start != -1)
    {
        if(!input[pos_start + label.length() + 1].isLetter())
        output.push_back(input.mid(pos_start, pos_end + ("</" + label + ">").size() - pos_start));
        
        ++index;
        pos_start = Find(input, ("<" + label), index + 1);
        pos_end = Find(input, ("</" + label + ">"), index + 1);

        if(pos_end == -1) pos_end = Find(input, ("/>"), index + 1);
    }
    return index;
}

int XmlOp::CountLabels(const QString xml_text, QString label)
{
    label = "<" + label;
    int cnt = 0;
    int search_pos = 0;
    int label_pos;
    while((label_pos = xml_text.indexOf(label, search_pos))!= -1)
    {
        if(!xml_text[label_pos + label.length()].isLetterOrNumber()) ++ cnt;
        
        search_pos = (label_pos + label.length());
    }

    return cnt;
}

bool XmlOp::ReplaceText(QString &xml_text, const QString &mark, const QString &replace)
{
    int pos = xml_text.indexOf(mark);
    if(pos == -1)
    {
        qDebug() << "mark " + mark + " not found";
        return false;
    }
    xml_text.replace(mark, replace);
    return true;
}

int XmlOp::AddRelationship(QString &xml_text, QString new_relationship)
{
    int pos_end_relationships = xml_text.indexOf("</Relationships>");
    if(pos_end_relationships == -1)
    {
        qDebug() << "Error: wrong xml content passed to function AddRelationship";
        return -1;
    }

    
    xml_text.insert(pos_end_relationships, new_relationship);
    return 1;
}

bool XmlOp::AddWorkBookSheet(QString &xml_text, int sheet_sn, int rid_sn)
{
    int pos_end_sheets = xml_text.indexOf("</sheets>");
    if(pos_end_sheets == -1)
    {
        qDebug() << "Error: wrong xml content passed to function AddWorkBookSheet";
        return false;
    }
    QString new_sheet = "<sheet name=\"Sheet${SHEET_SN}\" sheetId=\"${SHEET_SN}\" state=\"visible\" r:id=\"rId${RID_SN}\" />";
    
    int ret = ReplaceText(new_sheet, "${RID_SN}", QString::number(rid_sn));
    if(ret == false)
    {
        qDebug() << "Error: fail to replace RID_SN";
        return -1;
    }
    ret = ReplaceText(new_sheet, "${SHEET_SN}", QString::number(sheet_sn));
    if(ret == false)
    {
        qDebug() << "Error: fail to replace SHEET_SN";
        return -1;
    }
    xml_text.insert(pos_end_sheets, new_sheet);
    return true;
}

bool XmlOp::AddContentType(QString &xml_text, QString new_content_type)
{
    int pos_end_content_type = xml_text.indexOf("</Types>");
    if(pos_end_content_type == -1)
    {
        qDebug() << "Error: wrong xml content passed to function AddContentType";
        return false;
    }
    
    
    xml_text.insert(pos_end_content_type, new_content_type);
    return true;
}

bool XmlOp::AddText(QString &xml_text, QString new_text, QString label, int label_index)
{
    int pos = Find(xml_text, label, label_index);
    if(pos == -1)
    {
        qDebug() << "Can't find " << label << " ,with index " << label_index;
        return false;
    }
    xml_text.insert(pos, new_text);
    return true;
}

bool XmlOp::ModifyAttributes(QString &xml_text, QString xml_label, QString xml_attribute_label, QString xml_attribute_value, int xml_label_index)
{
    xml_label = "<" + xml_label;
    int xml_label_start_pos = Find(xml_text, xml_label, xml_label_index);
    if(xml_label_start_pos == -1)
    {
        return false;
    }
    int xml_label_end_pos = xml_label_start_pos + Find(xml_text.mid(xml_label_start_pos), ">", 1);

    int pos_start = -1;
    for(int i = xml_label_start_pos + xml_label.length(); i <= xml_label_end_pos; ++ i)
    {
        if((xml_text[i] != ' ' && xml_text[i] != '\n') && pos_start == -1) pos_start = i;
        if((xml_text[i] == ' ' || xml_text[i] == '\n' || xml_text[i] == '>') && pos_start != -1)
        {
            QString tmp_substr = xml_text.mid(pos_start, i - pos_start + 1);
            QString tmp_xml_attribute_label = tmp_substr.split('=')[0];
            if(tmp_xml_attribute_label == xml_attribute_label)
            {
                QString tmp_new_value;
                if(xml_attribute_value == "++")
                {
                    QStringList tmp_list = tmp_substr.split('\"');
                    bool isok = true;
                    int tmp_origin_value = tmp_list[1].toInt(&isok);
                    if(isok == false) return false;
                    tmp_new_value = QString::number(tmp_origin_value + 1);
                }
                else if(xml_attribute_value == "--")
                {
                    QStringList tmp_list = tmp_substr.split('\"');
                    bool isok = true;
                    int tmp_origin_value = tmp_list[1].toInt(&isok);
                    if(isok == false || tmp_origin_value == 0) return false;
                    tmp_new_value = QString::number(tmp_origin_value - 1);
                }
                else
                {
                    tmp_new_value = xml_attribute_value;
                }
                int quote_mark = -1;
                for(int j = pos_start; j < i; ++ j)
                {
                    if(xml_text[j] == '\"')
                    {
                        if(quote_mark == -1)
                        {
                            quote_mark = j;
                        }
                        else
                        {
                            xml_text.replace(quote_mark + 1, j - quote_mark - 1, tmp_new_value);
                            break;
                        }
                    }
                    
                }
                break;
            }
            pos_start = -1;
        }
    }
    return true;
}

QString XmlOp::GetAttribute(QString &xml_text, QString xml_label, QString xml_attribute_label, QString xml_feature_attribute_label)
{
    //qDebug() << "XmlOp::GetAttribute: xml label: "  << xml_label << " attribute_label: " << xml_attribute_label << " feature_attribute_label: " << xml_feature_attribute_label;
    int label_count = 0;
    int pos_start = Find(xml_text, "<" + xml_label, ++ label_count);
    int pos_end = 0;
    std::map<QString, QString> label_2_attribute_map;
    while(pos_start != -1)
    {
        if(xml_text[pos_start + xml_label.length() + 1] != ' ')
        {
            pos_start = Find(xml_text, "<" + xml_label, ++ label_count);
            continue;
        }
        pos_end = xml_text.indexOf(">", pos_start);
        pos_start += (2 + xml_label.length());
        QString sub_xml_label_text = xml_text.mid(pos_start, pos_end - pos_start - 1);
        if(sub_xml_label_text.contains(xml_feature_attribute_label))
        {
            QStringList labels = sub_xml_label_text.split(" ");
            QStringList tmp_labels;
            for(auto a : labels)
            {
                if(a.size() > 3)
                {
                    QString tmp_line;
                    for(int i = 0; i < a.size(); ++ i) if(a[i] != '\n') tmp_line += a[i];
                    tmp_labels.push_back(tmp_line);
                }
            }
            labels = tmp_labels;
            for(auto a : labels) 
            {
                QStringList label_2_attribute = a.split("=");
                label_2_attribute_map[label_2_attribute[0]] = label_2_attribute[1];
            }
            break;
        }
        pos_start = Find(xml_text, "<" + xml_label, ++ label_count);
    }
    
    if(label_2_attribute_map.empty() || label_2_attribute_map[xml_attribute_label].isEmpty()) return NULL;
    else
    {
        QString res = label_2_attribute_map[xml_attribute_label];
        RemoveOuterQuotation(res);
        return res;
    }
}

QString XmlOp::GetCellHeight(QString &xml_text, QString cell_sn)
{
    QString row_sn;
    for(int i = 0; i < cell_sn.length(); ++ i)
    {
        if(cell_sn[i].isDigit())
        {
            row_sn = cell_sn.mid(i);
            break;
        }
    }

    QString height = "12.8";
    std::vector<QString> rows_vector;
    int rows = AnalyzeXmlLabels(xml_text, rows_vector, "row");

    for(int i = rows - 1; i >= 0; -- i)
    {
        std::vector<QString> cell_vector;
        AnalyzeXmlLabels(rows_vector[i], cell_vector, "c");
        for(auto a : cell_vector)
        {
            
            if(a.contains(cell_sn))
            {
                int tmp_sub_str_pos = xml_text.indexOf(rows_vector[i]);
                QString tmp_sub_str = xml_text.mid(tmp_sub_str_pos);
                height = GetAttribute(rows_vector[i], "row", "ht", row_sn);
            }
        }
    }
    return height;
}

int XmlOp::CountStringSn(QString &xml_text, QString str)
{
    std::vector<QString> labels_list;
    bool ret;

    int counts = XmlOp::AnalyzeXmlLabels(xml_text, labels_list, "si");
    
    for(int i = 0; i < counts; i ++)
    {
        std::vector<QString> tmp_vector;
        XmlOp::AnalyzeXmlLabels(labels_list[i], tmp_vector, "t");
        
        XmlOp::ExtractTexts(tmp_vector);
        QString tmp_text;
        for(auto a : tmp_vector) tmp_text += a;
        if(tmp_text == str)
        {
            ret = XmlOp::ModifyAttributes(xml_text, "sst", "count", "++");
            if(ret == false)
            {
                qDebug() << "XmlOp::ModifyAttributes fail at increasing attribute count";
                return -1;
            }
            XmlOp::SaveXml(xml_text, xml_text);
            return i;
        }
    }

    int insert_pos = XmlOp::Find(xml_text, "</sst>", -1);
    QString string_model = "<si><t xml:space=\"preserve\">${TEXT}</t></si>";
    string_model.replace("${TEXT}", str);
    xml_text.insert(insert_pos, string_model);
    ret = XmlOp::ModifyAttributes(xml_text, "sst", "count", "++");
    if(ret == false)
    {
        qDebug() << "XmlOp::ModifyAttributes fail at increasing attribute count";
        return -1;
    }
    ret = XmlOp::ModifyAttributes(xml_text, "sst", "uniqueCount", "++");
    if(ret == false)
    {
        qDebug() << "XmlOp::ModifyAttributes fail at increasing attribute uniqueCount";
        return -1;
    }
    
    return counts;
}

int XmlOp::Find(QString src, QString des, int index)
{
    int _index = 1;
    int pos = 0;
    pos = src.indexOf(des);
    if(index == -1)
    {
        pos = src.lastIndexOf(des);
    }
    else
    {
        while(_index != index)
        {
            pos = src.indexOf(des, pos + des.size());
            ++ _index;
            if(pos == -1) break;
        }
    }
    
    return pos;
}

void XmlOp::ExtractTexts(std::vector<QString> &src)
{
    int f_pos, e_pos;
    QString label = "t";
    for(int i = 0; i < src.size(); ++ i)
    {
        QString tmp_text;
        int tmp_index = 0;

        f_pos = Find(src[i], "<" + label, ++ tmp_index);
        while(f_pos != -1)
        {
            for(; f_pos < src[i].size(); ++ f_pos) if(src[i][f_pos] == '>') break;
            f_pos ++;
            e_pos = Find(src[i], "</" + label + ">", tmp_index);
            tmp_text += src[i].mid(f_pos, e_pos - f_pos);
            
            f_pos = Find(src[i], "<" + label, ++ tmp_index);
        }
        src[i] = tmp_text;
    }
}

bool XmlOp::ReplaceSharedStringText(QString &xml_text, QString label, QString text)
{
    bool ret;
    ret = MergeTexts(xml_text, label);
    if(ret == false) return ret;
    ret = ReplaceText(xml_text, label, text);
    
    return ret;
}

std::map<QString, std::vector<QString>> XmlOp::GetCellSns(QString &xml_text, Info info, int direction)
{
    std::map<QString, std::vector<QString>> output;
    std::map<QString, std::vector<QString>> label_to_text_map;
    label_to_text_map = info.getLabelText();
    for(auto it = label_to_text_map.begin(); it != label_to_text_map.end(); ++ it)
    {
        QString tmp_label = it->first;
        std::vector<QString> tmp_texts = it->second;

        QString label_cell_sn = GetCellSn(xml_text, tmp_label);
        output[tmp_label].push_back(label_cell_sn);
        for(int i = 1; i < tmp_texts.size(); ++ i)
        {
            output[tmp_label].push_back(NextCellSn(output[tmp_label].back(), direction));
        }
    }

    return output;
}

void XmlOp::RemoveSpaces(QString &input)
{
    int pos_front, pos_back;
    for(int i = 0; i < input.length(); ++ i)
    {
        if(input[i] != ' ')
        {
            pos_front = i;
            break;
        }
    }

    for(int i = input.length() - 1; i >= 0; -- i)
    {
        if(input[i] != ' ')
        {
            pos_back = i;
            break;
        }
    }

    input = input.mid(pos_front, pos_back + 1 - pos_front);
}

void XmlOp::RemoveOuterQuotation(QString &input)
{
    int pos_front, pos_back;
    for(int i = 0; i < input.length(); ++ i)
    {
        if(input[i] != '"')
        {
            pos_front = i;
            break;
        }
    }

    for(int i = input.length() - 1; i >= 0; -- i)
    {
        if(input[i] != '"')
        {
            pos_back = i;
            break;
        }
    }

    input = input.mid(pos_front, pos_back + 1 - pos_front);
}

bool XmlOp::MergeTexts(QString &xml_text, QString &label)
{
    bool ret = true;

    std::vector<QString> shared_string_vector, text_vector;
    AnalyzeXmlLabels(xml_text, shared_string_vector, "si");
    
    for(auto a : shared_string_vector)
    {
        AnalyzeXmlLabels(a, text_vector, "t");
        ExtractTexts(text_vector);
        QString tmp_text;
        for(auto b : text_vector) tmp_text += b;
        ret =(ret == false ? false : ReplaceText(xml_text, a, "<si><t>" + tmp_text + "</t></si>"));
    }

    return ret;

}

QString XmlOp::GetCellSn(QString &xml_text, QString cell_text)
{
    std::vector<QString> row_vector, cell_vector;
    QString output = "";
    AnalyzeXmlLabels(xml_text, row_vector, "row");
    for(auto per_row : row_vector)
    {
        if(per_row.contains(">" + cell_text + "<"))
        {
            AnalyzeXmlLabels(per_row, cell_vector, "c");
            for(auto per_cell : cell_vector)
            {
                if(per_cell.contains(">" + cell_text + "<"))
                {
                    output = GetAttribute(per_cell, "c", "r", "\"");
                }
            }
        }
    }
    return output;
}

QString XmlOp::NextCellSn(QString current_cell_sn, int direction)
{
    QString new_cell_sn;
    if(direction == VERTICAL)
    {
        int pos = 0;
        for(int i = 0; i < current_cell_sn.size(); ++ i)
        {
            if(current_cell_sn[i].isDigit())
            {
                pos = i;
                break;
            }
        }
        QString front = current_cell_sn.mid(0, pos);
        QString rear = current_cell_sn.mid(pos);
        new_cell_sn = front + QString::number(rear.toInt() + 1);
        return new_cell_sn;
    }
    else
    {
        int pos = 0;
        for(int i = 0; i < current_cell_sn.size(); ++ i)
        {
            if(current_cell_sn[i].isDigit())
            {
                pos = i;
                break;
            }
        }
        QString front = current_cell_sn.mid(0, pos);
        QString rear = current_cell_sn.mid(pos);
        QString new_cell_sn;
        int addon = 1;
        int cur_index = front.size() - 1;
        QString alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        while(addon)
        {
            if(cur_index  == -1)
            {
                front = alphabet[-1 + addon] + front;
                addon = 0;
                continue;
            }
            int cur_num;
            for(int i = 0; i < 26; ++ i)
            {
                if((front[cur_index]) == alphabet[i])
                {
                    cur_num = i;
                    break;
                }
            }
            front[cur_index] = alphabet[(cur_num + addon) % 26];
            addon = (cur_num + addon) / 26;
            -- cur_index;
        }

        new_cell_sn = front + rear;
        return new_cell_sn;
    }
}