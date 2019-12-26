import psycopg2
from docxtpl import DocxTemplate
from win32com.client import Dispatch


def sql_fromDB(curs, sql_str):
    curs.execute(sql_str)
    result = curs.fetchall()
    return list(result[0])


def fetch_data(time_str):
    maxmin_sql = "SELECT MAX(\"Value\"),MIN(\"Value\") FROM \"public\".\"Ps_InSAR_huaibei\" WHERE \"Date\" = \'" + time_str + "\'"
    point_sql = "SELECT COUNT(\"Value\") FROM \"public\".\"Ps_InSAR_huaibei\" WHERE \"Date\" = '" + time_str + "'"
    point1_sql = "SELECT COUNT(\"Value\") FROM \"public\".\"Ps_InSAR_huaibei\" WHERE \"Date\" = '" + time_str + "' AND \"Value\" < -20"
    point2_sql = "SELECT COUNT(\"Value\") FROM \"public\".\"Ps_InSAR_huaibei\" WHERE \"Date\" = '" + time_str + "' AND \"Value\" < -10 AND \"Value\" > -20"
    point3_sql = "SELECT COUNT(\"Value\") FROM \"public\".\"Ps_InSAR_huaibei\" WHERE \"Date\" = '" + time_str + "' AND \"Value\" < 10 AND \"Value\" > -10"
    point4_sql = "SELECT COUNT(\"Value\") FROM \"public\".\"Ps_InSAR_huaibei\" WHERE \"Date\" = '" + time_str + "' AND \"Value\" < 20 AND \"Value\" > 10"
    point5_sql = "SELECT COUNT(\"Value\") FROM \"public\".\"Ps_InSAR_huaibei\" WHERE \"Date\" = '" + time_str + "' AND \"Value\" > 20"
    # 连接数据库获取写入报告中的数据
    conn = psycopg2.connect(database="postgres", user="postgres", password="valarmorghulis", host="127.0.0.1",
                            port="5432")
    curs = conn.cursor()
    max = sql_fromDB(curs, maxmin_sql)[0]  # 形变量最大值
    min = sql_fromDB(curs, maxmin_sql)[1]  # 形变量最小值
    point_num = sql_fromDB(curs, point_sql)[0]  # 所有点的总个数
    point1_num = sql_fromDB(curs, point1_sql)[0]  # 形变量<-20的点个数
    point2_num = sql_fromDB(curs, point2_sql)[0]  # 形变量>-20 & < -10的点个数
    point3_num = sql_fromDB(curs, point3_sql)[0]  # 形变量<10 & >-10的点个数
    point4_num = sql_fromDB(curs, point4_sql)[0]  # 形变量<20 & >10的点个数
    point5_num = sql_fromDB(curs, point5_sql)[0]  # 形变量>20的点个数
    # 结果保留小数点后三位
    max = round(max, 3)
    min = round(min, 3)
    point_num = round(point_num, 3)
    point1_num = round(point1_num, 3)
    point1_rate = round(point1_num / point_num * 100, 3)
    point2_num = round(point2_num, 3)
    point2_rate = round(point2_num / point_num * 100, 3)
    point3_num = round(point3_num, 3)
    point3_rate = round(point3_num / point_num * 100, 3)
    point4_num = round(point4_num, 3)
    point4_rate = round(point4_num / point_num * 100, 3)
    point5_num = round(point5_num, 3)
    point5_rate = round(point5_num / point_num * 100, 3)
    time_str_chinese = time_str[0:4] + "年" + time_str[4:6] + "月" + time_str[6:] + "日"
    #关闭数据库连接
    conn.close()
    data_dic = {
        'date': time_str_chinese,
        'set_max': max,
        'set_min': min,
        'point_num': point_num,
        'point1_num': point1_num,
        'point1_rate': point1_rate,
        'point2_num': point2_num,
        'point2_rate': point2_rate,
        'point3_num': point3_num,
        'point3_rate': point3_rate,
        'point4_num': point4_num,
        'point4_rate': point4_rate,
        'point5_num': point5_num,
        'point5_rate': point5_rate,
        'col_labels': ['-60至-20mm', '-20至-10mm', '-10至10mm', '10至20mm', '20至60mm'],
        'tbl_contents': [{'label': '监测点个数(个)',
                          'cols': [point1_num, point2_num, point3_num, point4_num, point5_num]},
                         {'label': '监测点占比（%）',
                          'cols': [point1_rate, point2_rate, point3_rate, point4_rate, point5_rate]}
                         ]
            }
    return data_dic


def export(time_str, pic1_path, pic2_path, out_doc_path):
    template_path = "./report/template.docx"
    context = fetch_data(time_str)
    tpl = DocxTemplate(template_path)
    tpl.replace_pic('pic1.jpg', pic1_path)
    tpl.replace_pic('pic2.jpg', pic2_path)
    tpl.render(context)
    tpl.save(out_doc_path)
    #word转PDF
    word = Dispatch('Word.Application')
    doc = word.Documents.Open(out_doc_path)
    doc.SaveAs(out_doc_path.replace('.docx', '.pdf'), FileFormat=17)
    doc.Close()
    word.Quit()


if __name__ == "__main__":
    time_str = "20191025"
    pic1_path = 'C:/Users/codhy/Desktop/Sar/output/20191025/20191025_shp.jpg'
    pic2_path = 'C:/Users/codhy/Desktop/Sar/output/20191025/20191025_rat.jpg'
    out_doc_path = 'C:/Users/codhy/Desktop/Sar/output/20191025/InSAR_report_20191025.docx'
    export(time_str, pic1_path, pic2_path, out_doc_path)