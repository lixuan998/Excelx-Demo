#include "excelop.h"
#include <ctime>
#include <sys/time.h>
#include <QApplication>
#include <QDebug>

void TestStyle(ExcelOp &op);
void TestRepacement(ExcelOp &op);
void TestWriteBatched(ExcelOp &op);
void TestDrawCell(ExcelOp &op);

#ifdef __APPLE__
    #define PATH "/Users/climatex/Documents"
    #define IMG_PATH "/Users/climatex/Downloads"
#elif __linux__
    #define PATH "/home/climatex/Documents/excel_test"
    #define IMG_PATH "/home/climatex/Pictures"
#endif

int main(int argc, char *argv[])
{
    QApplication app(argc, argv);
    //ExcelOp op1("/home/climatex/Documents/test_batch1.xlsx", "/home/climatex/Documents/test_batch1_output.xlsx");
    //ExcelOp op2("/home/climatex/Documents/test_batch2.xlsx", "/home/climatex/Documents/test_batch2_output.xlsx");
    QString path = PATH;
    ExcelOp op(path + "/empty.xlsx", path + "/empty_output.xlsx");
    //op1.open();
    //op2.open();
    op.open();
    //TestWriteBatched(op1);
    //TestWriteBatched(op2);

    //TestRepacement(op);
    TestDrawCell(op);
    //op1.close();
    //op2.close();
    op.close();
    return 0;
}

void TestStyle(ExcelOp &op)
{
    QString s1 = op.GetCellAttribute(1, "A1", "s");
    QString s2 = op.GetCellAttribute(1, "B1", "s");
    QString s3 = op.GetCellAttribute(1, "C1", "s");
    QString s4 = op.GetCellAttribute(1, "D1", "s");
    QString s5 = op.GetCellAttribute(1, "E1", "s");
    QString s6 = op.GetCellAttribute(1, "F1", "s");

    QString height1 = op.GetCellHeight(1, "A1");
    QString height2 = op.GetCellHeight(1, "B1");
    QString height3 = op.GetCellHeight(1, "C1");
    QString height4 = op.GetCellHeight(1, "D1");
    QString height5 = op.GetCellHeight(1, "E1");
    QString height6 = op.GetCellHeight(1, "F1");

    QString test_arr[]{"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"};
    QString cells[] = {"A", "B", "C", "D", "E", "F"};
    QString s[] = {s1, s2, s3, s4, s5, s6};

    QString heights[] = {height1, height2, height3, height4, height5, height6};
    for(int i = 0; i < 6; ++ i)
    {
        for(int j = 0; j < 26; ++ j)
        {
            op.WriteCell(1, cells[i] + QString::number(j + 10), test_arr[j], s[i], heights[i]);
        }
    }
}

void TestRepacement(ExcelOp &op)
{
    op.ReplaceSharedStringText("index", "这是");
    op.ReplaceSharedStringText("item1", "替换");
    op.ReplaceSharedStringText("item2", "文本");
    op.ReplaceSharedStringText("largeeeeeeeeeeeeeeeeeeeeeeeeitem", "的");
    op.ReplaceSharedStringText("item4", "测");
    op.ReplaceSharedStringText("item5", "试");
}

void TestWriteBatched(ExcelOp &op)
{
    Info info;
    QString data[4][5] = {{"${number1}", "A", "AA", "AAA", "AAAA"}, {"${number2}", "B", "BB", "BBB", "BBBB"}, {"${number3}", "8", "9", "10", "11"}, {"${number4}", "12", "13", "14", "15"}};
    for(int i = 0; i < 4; ++ i)
    {
        for(int j = 1; j < 5; ++ j)
        {
            op.AddInfo(info, data[i][0], data[i][j]);
        }
    }

    op.WriteBatch(1, info);
}

void TestDrawCell(ExcelOp &op)
{
    QString img_path = IMG_PATH;
    op.DrawCell(1, img_path + "/111.png", "0", "0", "0", "0", 500);
    op.DrawCell(1, img_path + "/111.png", "2", "2", "2", "2", 1000);
    op.DrawCell(1, img_path + "/111.png", "6", "6", "6", "6", 1500);
}