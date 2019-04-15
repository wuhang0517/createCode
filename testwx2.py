import xlrd
import os
# -*- coding: utf-8 -*-

###########################################################################
## Python code generated with wxFormBuilder (version Jun 17 2015)
## http://www.wxformbuilder.org/
##
## PLEASE DO "NOT" EDIT THIS FILE!
###########################################################################

import wx
import wx.xrc
import xlrd
import string
import os


###########################################################################
## Class MyFrame1
###########################################################################

class CreateProject(wx.Frame):

  def __init__(self, parent):
    wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=wx.EmptyString, pos=wx.DefaultPosition, size=wx.Size(798, 638),
                      style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

    self.SetSizeHintsSz(wx.DefaultSize, wx.DefaultSize)

    gSizer1 = wx.GridSizer(0, 2, 0, 0)

    self.m_staticText1 = wx.StaticText(self, wx.ID_ANY, u"excel路径", wx.DefaultPosition, wx.DefaultSize, 0)
    self.m_staticText1.Wrap(-1)
    gSizer1.Add(self.m_staticText1, 0, wx.ALL, 5)

    self.excel_path = wx.FilePickerCtrl(self, wx.ID_ANY, wx.EmptyString, u"Select a file", u"*.*",
                                        wx.DefaultPosition, wx.DefaultSize, wx.FLP_DEFAULT_STYLE)
    gSizer1.Add(self.excel_pa1th, 0, wx.ALL, 5)

    self.m_staticText3 = wx.StaticText(self, wx.ID_ANY, u"包名", wx.DefaultPosition, wx.DefaultSize, 0)
    self.m_staticText3.Wrap(-1)
    gSizer1.Add(self.m_staticText3, 0, wx.ALL, 5)

    self.package_name = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
    gSizer1.Add(self.package_name, 0, wx.ALL, 5)

    self.m_staticText4 = wx.StaticText(self, wx.ID_ANY, u"生产文件类型", wx.DefaultPosition, wx.DefaultSize, 0)
    self.m_staticText4.Wrap(-1)
    gSizer1.Add(self.m_staticText4, 0, wx.ALL, 5)

    bSizer1 = wx.GridSizer(0, 2, 0, 0)

    self.project = wx.CheckBox(self, wx.ID_ANY, u"project", wx.DefaultPosition, wx.DefaultSize, 0)
    bSizer1.Add(self.project, 0, wx.ALL, 5)

    self.project_name = wx.TextCtrl(self, wx.ID_ANY, "", wx.DefaultPosition, wx.DefaultSize, 0)
    bSizer1.Add(self.project_name, 0, wx.ALL, 5)

    # bSizer1.AddSpacer(5)

    self.pojo = wx.CheckBox(self, wx.ID_ANY, u"pojo", wx.DefaultPosition, wx.DefaultSize, 0)
    bSizer1.Add(self.pojo, 0, wx.ALL, 5)

    self.pojo_path = wx.TextCtrl(self, wx.ID_ANY, "", wx.DefaultPosition, wx.DefaultSize, 0)
    bSizer1.Add(self.pojo_path, 0, wx.ALL, 5)

    self.mapper = wx.CheckBox(self, wx.ID_ANY, u"mapper", wx.DefaultPosition, wx.DefaultSize, 0)
    bSizer1.Add(self.mapper, 0, wx.ALL, 5)

    self.mapper_path = wx.TextCtrl(self, wx.ID_ANY, "", wx.DefaultPosition, wx.DefaultSize, 0)
    bSizer1.Add(self.mapper_path, 0, wx.ALL, 5)

    self.dao = wx.CheckBox(self, wx.ID_ANY, u"dao", wx.DefaultPosition, wx.DefaultSize, 0)
    bSizer1.Add(self.dao, 0, wx.ALL, 5)

    self.dao_path = wx.TextCtrl(self, wx.ID_ANY, "", wx.DefaultPosition, wx.DefaultSize, 0)
    bSizer1.Add(self.dao_path, 0, wx.ALL, 5)

    self.sql = wx.CheckBox(self, wx.ID_ANY, u"sql", wx.DefaultPosition, wx.DefaultSize, 0)
    bSizer1.Add(self.sql, 0, wx.ALL, 5)

    # 是否需要注释
    self.sql_commont = wx.CheckBox(self, wx.ID_ANY, u"是否需要注释", wx.DefaultPosition, wx.DefaultSize, 0)
    bSizer1.Add(self.sql_commont, 0, wx.ALL, 5)

    gSizer1.Add(bSizer1, 1, wx.EXPAND, 5)

    self.create_button = wx.Button(self, wx.ID_ANY, u"生成", wx.DefaultPosition, wx.DefaultSize, 0)
    gSizer1.Add(self.create_button, 0, wx.ALL, 5)

    self.SetSizer(gSizer1)
    self.Layout()

    self.Centre(wx.BOTH)

    # Connect Events
    self.create_button.Bind(wx.EVT_BUTTON, self.create_button_click)

    self.project.Bind(wx.EVT_CHECKBOX, self.project_check)

    self.pojo.Bind(wx.EVT_CHECKBOX, self.type_check)

    self.dao.Bind(wx.EVT_CHECKBOX, self.type_check)

    self.mapper.Bind(wx.EVT_CHECKBOX, self.type_check)

    self.sql.Bind(wx.EVT_CHECKBOX, self.type_check)

  def __del__(self):
    pass

  def create_button_click(self, event):
    event.Skip()

  def project_check(self, event):
    if self.project.GetValue() == True:
      self.pojo.SetValue(False)
      self.mapper.SetValue(False)
      self.dao.SetValue(False)
      self.sql.SetValue(False)
      self.project_name.SetValue(u"项目名称")
      self.project_name.SetEditable(True)
    else:
      self.project_name.SetValue("")
      self.project_name.SetEditable(False)

  def type_check(self, event):
    if self.pojo.GetValue() == True or self.dao.GetValue() == True or self.mapper.GetValue() == True or self.sql.GetValue() == True:
      self.project.SetValue(False)
      self.project_name.SetValue(u"")
      self.project_name.SetEditable(False)


class myapp(CreateProject):

  def create_button_click(self, event):
    # 获取excel路径
    app_excel_path = self.excel_path.GetPath()

    path = str(app_excel_path)
    file_name = path[path.rfind('\\') + 1:path.find('.')]
    if app_excel_path.rfind('.') == -1:
      pass
    # 读取excel
    app_excel = xlrd.open_workbook(self.excel_path.GetPath())
    # 获取sheet页
    class_sheet = app_excel.sheet_by_index(0)

    app_pojo_path = str(self.pojo_path.GetValue()).replace(".", "\\\\")

    app_mapper_path = str(self.mapper_path.GetValue()).replace(".", "\\\\")

    app_dao_path = str(self.dao_path.GetValue()).replace(".", "\\\\")

    app_sql_path = ""
    if self.pojo.GetValue() == True:
      if not os.path.exists(app_pojo_path):
        os.makedirs(app_pojo_path)
      self.create_pojo(class_sheet, str(self.pojo_path.GetValue()))

    if self.mapper.GetValue() == True:
      if not os.path.exists(app_mapper_path):
        os.makedirs(app_mapper_path)
      self.create_mapper(class_sheet, str(self.mapper_path.GetValue()), str(self.pojo_path.GetValue()),
                         str(self.dao_path.GetValue()))

    if self.dao.GetValue() == True:
      if not os.path.exists(app_dao_path):
        os.makedirs(app_dao_path)
      self.crteate_dao(class_sheet, str(self.dao_path.GetValue()))

    if self.sql.GetValue() == True:
      is_sql_comment = self.sql_commont.GetValue()
      app_sql_path = "sql"
      if not os.path.exists(app_sql_path):
        os.makedirs(app_sql_path)
      self.create_sql(class_sheet, app_sql_path + "/" + file_name + ".sql", is_sql_comment)
    wx.MessageBox("生成成功", "MESSAGE", wx.OK | wx.ICON_INFORMATION)

  def crteate_dao(self, sheet, path):
    all_class_names = []
    for i in range(0, sheet.nrows):
      if str(sheet.row_values(i)[1]) is not '' and str(sheet.row_values(i)[0]) is '' and not str(sheet.row_values(i)[1]).startswith('IX', 0, 2):
        class_str_old = string.capwords(str(sheet.row_values(i)[1]).replace('_', ' '))
        class_str_new = class_str_old.replace(' ', '')
        all_class_names.append(class_str_new)
    for class_name in all_class_names:
      class_dao_str = []
      class_name_lower = class_name[0].lower() + class_name[1:]
      file_name = open(path.replace(".", "\\\\") + "/" + class_name + "Mapper.java", "w", encoding="UTF-8")
      class_dao_str.append("package " + path + ";\n\n")
      class_dao_str.append("import java.util.List;\n")
      class_dao_str.append("import " + path + ".entity." + class_name + ";\n\n")
      class_dao_str.append("public interface " + class_name + "Dao { \n")
      class_dao_str.append("\t/**\n\t * 增加一条\n\t * param +"+class_name_lower+"+\n\t */\n")
      class_dao_str.append("\tvoid insert" + class_name + " (" + class_name + " " + class_name_lower + ");\n")
      class_dao_str.append("\t/**\n\t * 修改一条\n\t * param +"+class_name_lower+"+\n\t */\n")
      class_dao_str.append("\tvoid update" + class_name + " (" + class_name + " " + class_name_lower + ");\n")
      class_dao_str.append("\t/**\n\t * 删除一条\n\t * param +"+class_name_lower+"+\n\t */\n")
      class_dao_str.append("\tvoid delete" + class_name + " (" + class_name + " " + class_name_lower + ");\n")
      class_dao_str.append("\t/**\n\t * 增加多条\n\t * param +"+class_name_lower+"+\n\t */\n")
      class_dao_str.append(
        "\tvoid insert" + class_name + "s (List<" + class_name + "> " + class_name_lower + "s);\n")
      class_dao_str.append("\t/**\n\t * 修改多条\n\t * param +"+class_name_lower+"+\n\t */\n")
      class_dao_str.append(
        "\tvoid update" + class_name + "s (List<" + class_name + "> " + class_name_lower + "s);\n")
      class_dao_str.append("\t/**\n\t * 通过主建查询\n\t * param +"+class_name_lower+"+\n\t */\n")
      class_dao_str.append(
        "\tList<" + class_name + "> select" + class_name + "(" + class_name + " " + class_name_lower + ");\n")
      class_dao_str.append("}\t")
      file_name.write("".join(class_dao_str))
      file_name.close()

  def create_mapper(self, sheet, mapper_path, pojo_path, dao_path):
    rows_num = 1
    # mapper开头拼接
    # = []
    while rows_num < sheet.nrows:
      if sheet.row_values(rows_num)[1] is '' or str(sheet.row_values(rows_num)[1]).startswith('IX', 0, 2):
        rows_num += 1
        continue
      cl_class_old = str(sheet.row_values(rows_num)[1])
      cl_class_new = string.capwords(cl_class_old.replace('_', ' '))
      class_name_str = cl_class_new.replace(' ', '')
      table_name_str = str(sheet.row_values(rows_num)[1])
      mapper_file = open(mapper_path.replace(".", "\\\\") + "/" + class_name_str + "Mapper.xml", 'w', encoding="utf-8")
      param_map = {}

      map_start = []
      map_start.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
      map_start.append(
        "<!DOCTYPE mapper PUBLIC \"-//mybatis.org//DTD Mapper 3.0//EN\" \"http://mybatis.org/dtd/mybatis-3-mapper.dtd\" >\n")
      map_start.append(
        "<mapper namespace=\"" + dao_path + "." + class_name_str + "Mapper\">\n")
      # insert 拼接
      insert_sql_one = []
      insert_sql_list = []
      insert_sql_one_middle=[]
      insert_sql_list_middle=[]
      insert_sql_one.append(
        "<insert id=\"insert" + class_name_str + "\" parameterType=\"" + pojo_path + "." + class_name_str
        + "\">\n")
      insert_sql_list.append("<insert id=\"insert" + class_name_str + "s\" parameterType=\"java.util.List\">\n")
      insert_sql_one.append("\tINSERT INTO " + table_name_str + " (\n")
      insert_sql_list.append("\tINSERT INTO " + table_name_str + " (\n")

      # select拼接
      select_sql = []
      select_sql.append(
        "<select id=\"select" + class_name_str + "\" parameterType=\"" + pojo_path + "." + class_name_str
        + "\"\n\tresultType=\"java.util.List\">\n\tSELECT \n")

      # update拼接
      update_sql_one = []
      update_sql_list = []
      update_sql_list.append("<update id=\"update" + class_name_str + "s\" parameterType=\"java.util.List\">\n")
      update_sql_list.append("\tBEGIN\n")
      update_sql_list.append("\t<foreach collection=\"list\" separator=\";\" item=\"item\" index=\"index\">\n")
      update_sql_list.append("\t\tUPDATE " + table_name_str + " SET \n")
      update_sql_one.append("<update id=\"update" + class_name_str + "\" parameterType=\"" + pojo_path + "." + class_name_str+ "\">\n")
      update_sql_one.append("\t\tUPDATE " + table_name_str + " SET \n")
      update_sql_one.append("\t\tWHERE\n")
      update_sql_list.append("\t\tWHERE\n")

      # delete拼接
      delete_sql = []
      delete_sql.append("<delete id=\"delete" + class_name_str + "\" parameterType=\"" + pojo_path + "." + class_name_str+ "\">\n")
      delete_sql.append("\tdelete from " + table_name_str + " where \n")
      delete_sql.append("\t<foreach collection=\"list\" separator=\"or\" item=\"item\" index=\"index\">\n")

      for i in range(rows_num, sheet.nrows):
        rows_num += 1
        if sheet.row_values(i)[1] is '':
          # property_str = []
          break
        if sheet.row_values(i)[3] is '':
          continue
        cl_db = str(sheet.row_values(i)[1])
        if i == sheet.nrows - 1 or sheet.row_values(i + 1)[1] is '':
          insert_sql_one.append("\t\t" + cl_db + "\n")
          insert_sql_list.append("\t\t" + cl_db + "\n")
          select_sql.append("\t\t" + cl_db + "\n")
        else:
          insert_sql_one.append("\t\t" + cl_db + ",\n")
          insert_sql_list.append("\t\t" + cl_db + ",\n")
          select_sql.append("\t\t" + cl_db + ",\n")

        cl_class_old = str(sheet.row_values(i)[1])
        cl_class_new = string.capwords(cl_class_old.replace('_', ' '))
        cl_str_property = cl_class_new.replace(' ', '')
        cl_str = cl_str_property[0].lower() + cl_str_property[1:]
        cl_type = ""
        if i == 0:
          new_str = cl_str_property
          continue
        # insert拼接
        if "VARCHAR" in str(sheet.row_values(i)[3]):
          cl_type = "VARCHAR"
        elif "DATE" in str(sheet.row_values(i)[3]):
          cl_type = "DATE"
        elif "NUMBER" in str(sheet.row_values(i)[3]) or "DECIMAL" in str(sheet.row_values(i)[3]):
          cl_type = "DOUBLE"
        if i == sheet.nrows - 1 or sheet.row_values(i + 1)[1] is '':
          if "LAST_UPDATE_TIME" == str(sheet.row_values(i)[1]) or "CREATE_DATE" == str(sheet.row_values(i)[1]):
            insert_sql_one_middle.append("\t\tSYSDATE\n")
            insert_sql_list_middle.append("\t\tSYSDATE\n")
          else:
            insert_sql_one_middle.append("\t\t#{" + cl_str + ", jdbcType=" + cl_type + "}\n")
            insert_sql_list_middle.append("\t\t#{item." + cl_str + ", jdbcType=" + cl_type + "}\n")
        else:
          if "LAST_UPDATE_TIME" == str(sheet.row_values(i)[1]) or "CREATE_DATE" == str(sheet.row_values(i)[1]):
            insert_sql_one_middle.append("\t\tSYSDATE,\n")
            insert_sql_list_middle.append("\t\tSYSDATE,\n")
          elif str(sheet.row_values(i)[5]) == "N":
            insert_sql_one_middle.append("\t\t#{" + cl_str + "},\n")
            insert_sql_list_middle.append("\t\t#{item." + cl_str + "},\n")
          else:
            insert_sql_one_middle.append("\t\t#{" + cl_str + ", jdbcType=" + cl_type + "},\n")
            insert_sql_list_middle.append("\t\t#{item." + cl_str + ", jdbcType=" + cl_type + "},\n")

        # update拼接
        if not str(sheet.row_values(i)[4]) == "":
          param_map[cl_class_old] = cl_str
          continue

        if "VERSION" == cl_class_old:
          update_sql_one.append('\t\t' + cl_class_old + '=' + cl_class_old + '+1,\n')
          update_sql_list.append('\t\t' + cl_class_old + '=' + cl_class_old + '+1,\n')
        elif "LAST_UPDATE_TIME" == cl_class_old:
          update_sql_one.append('\t\tSYSDATE,\n')
          update_sql_list.append('\t\tSYSDATE,\n')
        else:
          update_sql_one.append("\t\t<if test=\"" + cl_str + "!=null and " + cl_str + " !=''\">\n")
          update_sql_list.append("\t\t<if test=\"item." + cl_str + "!=null and item." + cl_str + " !=''\">\n")
        if i == sheet.nrows - 1 or sheet.row_values(i + 1)[1] is '':
          update_sql_one.append("\t\t\t" + cl_class_old + "=" + "#{" + cl_str + "}\n")
          update_sql_list.append("\t\t\t" + cl_class_old + "=" + "#{item." + cl_str + "}\n")
        else:
          update_sql_one.append("\t\t\t" + cl_class_old + "=" + "#{" + cl_str + "},\n")
          update_sql_list.append("\t\t\t" + cl_class_old + "=" + "#{item." + cl_str + "},\n")
        update_sql_one.append("\t\t</if>\n")
        update_sql_list.append("\t\t</if>\n")

      insert_sql_list.append("\t <foreach item=\"item\" collection=\"list\" index=\"index\" separator=\"union all \">\n")
      insert_sql_list.append("\tSELECT \n")
      insert_sql_list.append("".join(insert_sql_list_middle))
      insert_sql_list.append("\tfrom dual \n")
      insert_sql_list.append("\t</foreach>\n")
      insert_sql_list.append("\t)\n")
      insert_sql_list.append("</insert>\n")
      insert_sql_one.append("".join(insert_sql_one_middle))
      insert_sql_one.append("\t)\n")
      insert_sql_one.append("</insert>\n")
      select_sql.append("\tFROM " + table_name_str + " WHERE \n")

      i = 0
      for key, value in param_map.items():
        i += 1
        if i == len(param_map):
          update_sql_one.append("\t\t" + key + "=#{" + value + "}\n")
          update_sql_list.append("\t\t" + key + "=#{item." + value + "}\n")
          delete_sql.append("\t\t" + key + "=#{item." + value + "}\n")
          select_sql.append("\t\t" + key + "=#{" + value + "}\n")
          break
        update_sql_one.append("\t\t" + key + "=#{" + value + "} and \n")
        update_sql_list.append("\t\t" + key + "=#{item." + value + "} and \n")
        delete_sql.append("\t\t(" + key + "=#{item." + value + "} and \n")
        select_sql.append("\t\t" + key + "=#{" + value + "} and \n")

      update_sql_list.append("\t</foreach>\n")
      update_sql_list.append("\t;END;\n")
      update_sql_list.append("</update>\n")
      update_sql_one.append("</update>\n")
      delete_sql.append("\t</foreach>\n</delete>\n")
      select_sql.append("</select>")
      mapper_str = map_start + insert_sql_one + insert_sql_list + update_sql_one + update_sql_list + delete_sql + select_sql
      mapper_str.append("\n</mapper>")
      mapper_file.write("".join(mapper_str))
      mapper_file.close()

  def create_pojo(self, sheet, path):
    rows_num = 1
    pojo_code = {}
    class_name_num = []
    while rows_num < sheet.nrows:
      type_parm = set()
      property_str = []
      new_str = ''
      property_str.append("package " + path + ";\n\n")
      property_str.append("import javax.validation.constraints.NotNull;\n")
      property_str.append("import com.fasterxml.jackson.annotation.JsonProperty;\n")
      property_str.append("import lombok.Data;\n")
      property_str.append("import lombok.AllArgsConstructor;\n")
      property_str.append("import lombok.Getter;\n")
      property_str.append("import lombok.NoArgsConstructor;\n")
      property_str.append("import lombok.Setter;\n\n")
      property_str.append("@Data\n")
      property_str.append("@AllArgsConstructor\n")
      property_str.append("@NoArgsConstructor\n")
      for i in range(rows_num, sheet.nrows):
        rows_num += 1
        if sheet.row_values(i)[1] is '' or str(sheet.row_values(i)[1])[:2] == 'IX':
          class_name_num.append(i + 1)
          # property_str = []
          break
        cl_class_old = str(sheet.row_values(i)[1])
        cl_class_new = string.capwords(cl_class_old.replace('_', ' '))
        cl_str = cl_class_new.replace(' ', '')
        if i in class_name_num or i == 1:
          new_str = cl_str
          str1 = "public class " + cl_str + " { \n"
          property_str.append(str1)
          continue
        cl_str = cl_str[0].lower() + cl_str[1:]
        property_str.append("\t/**\n\t * " + str(sheet.row_values(i)[2]) + "\n" + "\t */\n")
        if str(sheet.row_values(i)[5]) == "N":
          str2 = "\t@NotNull\n"
          property_str.append(str2)
        if cl_class_old[1] == "_":
          type_parm.add("json")
          json_str = "@Getter(onMethod=@__(@JsonProperty(\"" + cl_str + "\")))\n"
          json_str = json_str + "\t@Setter(onMethod=@__(@JsonProperty(\"" + cl_str + "\")))\n"
          property_str.append("\t" + json_str)
        if "VARCHAR" in str(sheet.row_values(i)[3]):
          private_str = "private String "
        elif "DATE" in str(sheet.row_values(i)[3]):
          type_parm.add("localdate")
          private_str = "private LocalDateTime "
        elif "," in str(sheet.row_values(i)[3]):
          private_str = "private Double "
        else:
          if cl_str == 'version':
            private_str = "private Long "
          else:
            private_str = "private Integer "
        property_str.append("\t" + private_str + cl_str + ";\n")
      if property_str == []:
        pass
      else:
        property_str.append("}\n")
        if "localdate" in type_parm:
          property_str.insert(2, "import java.time.LocalDateTime;\n")
        pojo_code[new_str + '.java'] = ''.join(property_str)
    for (key, value) in pojo_code.items():
      if key ==".java":
        continue
      with open(path.replace(".", "\\\\") + "/" + key, 'w', encoding='utf-8') as wfile:
        wfile.write(value)

  def create_sql(self, sheet, path, is_sql_comment):
    row_num = 1
    CREATE_SQL = []
    PK_SQL = []
    COMMENT_SQL = []
    table_name = ''
    column_name = ''
    with open(path, 'w', encoding='utf-8') as wfile:
      for i in range(row_num, sheet.nrows):
        # row_num += 1
        if sheet.row_values(i)[0] == '':
          if str(sheet.row_values(i)[1]) == '' or str(sheet.row_values(i)[1]).startswith('IX', 0, 2):
            continue
          table_name = str(sheet.row_values(i)[1])
          PK_SQL = []
          COMMENT_SQL = []
          CREATE_SQL.append('CREATE TABLE ')
          CREATE_SQL.append(table_name)
          CREATE_SQL.append('\n(\n')
          PK_SQL.append('ALTER TABLE ')
          PK_SQL.append(table_name)
          PK_SQL.append(' ADD CONSTRAINT ')
          PK_SQL.append('PK_')
          PK_SQL.append(table_name)
          PK_SQL.append(' PRIMARY KEY (')

          COMMENT_SQL.append("COMMENT ON TABLE ")
          COMMENT_SQL.append(table_name)
          COMMENT_SQL.append(" IS ")
          COMMENT_SQL.append("'")
          COMMENT_SQL.append(str(sheet.row_values(i)[2]))
          COMMENT_SQL.append("';\n")
          continue
        column_name = str(sheet.row_values(i)[1])
        CREATE_SQL.append("\t")
        CREATE_SQL.append(column_name)
        CREATE_SQL.append(' ' * (40 - column_name.__len__()))
        number_str = ''
        if str(sheet.row_values(i)[3]).startswith('DECIMAL'):
          number_str = 'NUMBER' + str(sheet.row_values(i)[3])[7:]
        else:
          number_str = str(sheet.row_values(i)[3])
        CREATE_SQL.append(number_str)
        if sheet.row_values(i)[5] is not '':
          CREATE_SQL.append(' ' * (20 - number_str.__len__()))
          CREATE_SQL.append('NOT NULL')
        CREATE_SQL.append(',')
        CREATE_SQL.append('\n')
        # 列注释
        COMMENT_SQL.append("COMMENT ON COLUMN ")
        COMMENT_SQL.append(table_name)
        COMMENT_SQL.append(".")
        COMMENT_SQL.append(column_name)
        COMMENT_SQL.append(" IS ")
        COMMENT_SQL.append("'")
        COMMENT_SQL.append(str(sheet.row_values(i)[2]))
        COMMENT_SQL.append("';\n")
        if str(sheet.row_values(i)[4]) is not '':
          PK_SQL.append(column_name)
          PK_SQL.append(',')
        if i >= sheet.nrows - 1 or str(sheet.row_values(i + 1)[1]) is '':
          del CREATE_SQL[-2]
          del PK_SQL[-1]
          PK_SQL.append(');\n')
          CREATE_SQL.append(');\n')
          COMMENT_SQL.append("\n")
          if is_sql_comment == False:
            COMMENT_SQL = []
          CREATE_SQL = CREATE_SQL + PK_SQL + COMMENT_SQL
      wfile.write("".join(CREATE_SQL))


if __name__ == '__main__':
  app = wx.App()

  main_win = myapp(None)
  main_win.Show()
  app.MainLoop()
