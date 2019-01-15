import xlrd
import string
import os

rows_num = 0
rows_num_start = 0


# 项目生成器
def main():
  # 读取excel
  excel_name = xlrd.open_workbook(r'项目excel.xlsx')
  # 读取标签页
  sheets = excel_name.sheet_names()
  sheet_pom = excel_name.sheet_by_name(r"pom")
  path_list = make_dirs(sheet_pom)
  print("\n".join(path_list))
  sheet_class = excel_name.sheet_by_name(r"class")
  # 读取excel总共有几行
  sheet_rows_num = sheet_pom.nrows
  # 读取excel总共有几列
  sheet_cols_num = sheet_pom.ncols
  # 创建目录
  # 目录分割符
  seprarator = """\\"""
  # pom文件内容
  write_pom(sheet_pom, path_list[0] + seprarator + "pom.xml")
  # java代码
  # path_list = ['tmc-risk',
  #                 'tmc-risk\\src\\main\\java\\cn\\gov\\costoms\\h2018\\tmc\\trd\\pojo',
  #                 'tmc-risk\\src\\main\\java\\cn\\gov\\costoms\\h2018\\tmc\\trd\\dao',
  #                 'tmc-risk\\src\\main\\java\\cn\\gov\\costoms\\h2018\\tmc\\trd\\service',
  #                 'tmc-risk\\src\\main\\java\\cn\\gov\\costoms\\h2018\\tmc\\trd\\service\\impl',
  #                 'tmc-risk\\src\\main\\java\\cn\\gov\\costoms\\h2018\\tmc\\trd\\controller',
  #                 'tmc-risk\\src\\main\\java\\cn\\gov\\costoms\\h2018\\tmc\\trd\\config',
  #                 'tmc-risk\\src\\main\\resources\\mapper',
  #                 'tmc-risk\\src\\test\\java\\cn\\gov\\costoms\\h2018\\tmc\\trd\\pojo',
  #                 'tmc-risk\\src\\test\\java\\cn\\gov\\costoms\\h2018\\tmc\\trd\\dao',
  #                 'tmc-risk\\src\\test\\java\\cn\\gov\\costoms\\h2018\\tmc\\trd\\service',
  #                 'tmc-risk\\src\\test\\java\\cn\\gov\\costoms\\h2018\\tmc\\trd\\service\\impl',
  #                 'tmc-risk\\src\\test\\java\\cn\\gov\\costoms\\h2018\\tmc\\trd\\controller',
  #                 'tmc-risk\\src\\test\\java\\cn\\gov\\costoms\\h2018\\tmc\\trd\\config',
  #                 'tmc-risk\\src\\main\\resources',
  #                 'tmc-risk\\src\\main\\java\\cn\\gov\\costoms\\h2018\\tmc\\trd']
  class_and_classstr = java_code(sheet_class, path_list[1])
  class_code = class_and_classstr[0]
  class_num = class_and_classstr[1]
  class_strs = []
  for cl_num in class_num:
    class_strs.append(str(sheet_class.row_values(cl_num)[1]))
  # del class_code[".java"]

  write_pojo(class_code, seprarator, path_list)
  write_dao(tuple(class_code.keys()), seprarator, path_list)
  write_service(tuple(class_code.keys()), seprarator, path_list)
  write_service_impl(tuple(class_code.keys()), seprarator, path_list)
  write_controller(tuple(class_code.keys()), seprarator, class_strs, path_list)
  write_mapper(sheet_class, seprarator, path_list)
  write_starter(seprarator, path_list[-1])
  write_test_service_impl(tuple(class_code.keys()), seprarator, path_list)
  write_test_controller(tuple(class_code.keys()), seprarator, path_list)
  write_application(path_list[-2], seprarator)


# java代码处理
def java_code(sheet, path):
  global rows_num
  global rows_num_start
  pojo_code = {}
  class_name_num = []
  # print(".".join(path.split("""\\""")[4:]))
  while rows_num < sheet.nrows:
    type_parm = set()
    property_str = []
    new_str = ''
    property_str.append("package " + ".".join(path.split("""\\""")[4:]) + ";\n\n")
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
      if sheet.row_values(i)[1] is '':
        class_name_num.append(i + 1)
        # property_str = []
        break
      cl_class_old = str(sheet.row_values(i)[1])
      cl_class_new = string.capwords(cl_class_old.replace('_', ' '))
      cl_str = cl_class_new.replace(' ', '')
      if i in class_name_num or i == 0:
        new_str = cl_str
        str1 = "public class " + cl_str + " { \n"
        property_str.append(str1)
        continue
      cl_str = cl_str[0].lower() + cl_str[1:]
      property_str.append("\t/**\n\t\0*" + str(sheet.row_values(i)[2]) + "\n" + "\t\0*/\n")
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
  rows_num = rows_num_start
  return pojo_code, class_name_num


# pojo
def write_pojo(class_code, seprarator, path_list):
  # 循环创建文件
  for key, value in class_code.items():
    class_file = open(path_list[1] + seprarator + key, 'w', encoding="utf-8")
    class_file.write(value)
    class_file.close()


# dao
def write_dao(class_names, seprarator, path_list):
  for class_name in class_names:
    class_name = class_name[:-5]
    class_dao_str = []
    class_name_lower = class_name[0].lower() + class_name[1:]
    file_name = open(path_list[2] + seprarator + class_name + "Dao.java", "w", encoding="UTF-8")
    class_dao_str.append("package " + ".".join(path_list[2].split("""\\""")[4:]) + ";\n\n")
    class_dao_str.append("import java.util.List;\n")
    class_dao_str.append("import " + ".".join(path_list[1].split("""\\""")[4:]) + "." + class_name + ";\n\n")
    class_dao_str.append("public interface " + class_name + "Dao { \n")
    class_dao_str.append("\t /** \n \t\0 * 增加一条 \n \t\0\0*/\n")
    class_dao_str.append("\t void insert" + class_name + " (" + class_name + " " + class_name_lower + ");\n")
    class_dao_str.append("\t /** \n \t\0 * 修改一条 \n \t\0\0*/\n")
    class_dao_str.append("\t void update" + class_name + " (" + class_name + " " + class_name_lower + ");\n")
    class_dao_str.append("\t /** \n \t\0 * 删除一条 \n \t\0\0*/\n")
    class_dao_str.append("\t void delete" + class_name + " (" + class_name + " " + class_name_lower + ");\n")
    class_dao_str.append("\t /** \n \t\0 * 增加多条 \n \t\0\0*/\n")
    class_dao_str.append(
      "\t void insert" + class_name + "s (List<" + class_name + "> " + class_name_lower + "s);\n")
    class_dao_str.append("\t /** \n \t\0 * 修改多条 \n \t\0\0*/\n")
    class_dao_str.append(
      "\t void update" + class_name + "s (List<" + class_name + "> " + class_name_lower + "s);\n")
    class_dao_str.append("\t /** \n \t\0 * 通过主建查询 \n \t\0\0*/\n")
    class_dao_str.append(
      "\t List<" + class_name + "> select" + class_name + "(" + class_name + " " + class_name_lower + ");\n")
    class_dao_str.append("}\t")
    file_name.write("".join(class_dao_str))
    file_name.close()


# service
def write_service(class_names, seprarator, path_list):
  for class_name in class_names:
    class_name = class_name[:-5]
    class_service_str = []
    class_name_lower = class_name[0].lower() + class_name[1:]
    file_name = open(path_list[3] + seprarator + class_name + "Service.java", "w", encoding="UTF-8")
    class_service_str.append("package " + ".".join(path_list[3].split("""\\""")[4:]) + ";\n\n")
    class_service_str.append("import java.util.List;\n")
    class_service_str.append("import " + ".".join(path_list[1].split("""\\""")[4:]) + "." + class_name + ";\n\n")
    class_service_str.append("public interface " + class_name + "Service { \n")
    class_service_str.append("\t /** \n \t\0 * 增加一条 \n \t\0\0*/\n")
    class_service_str.append("\tString insert" + class_name + " (" + class_name + " " + class_name_lower + ");\n")
    class_service_str.append("\t /** \n \t\0 * 修改一条 \n \t\0*/\n")
    class_service_str.append("\tString update" + class_name + " (" + class_name + " " + class_name_lower + ");\n")
    class_service_str.append("\t /** \n \t\0 * 删除一条 \n \t\0\0*/\n")
    class_service_str.append("\tString delete" + class_name + " (" + class_name + " " + class_name_lower + ");\n")
    class_service_str.append("\t /** \n \t\0 * 增加多条 \n \t\0\0*/\n")
    class_service_str.append(
      "\tString insert" + class_name + "s (List<" + class_name + "> " + class_name_lower + "s);\n")
    class_service_str.append("\t /** \n \t\0 * 修改多条 \n \t\0\0*/\n")
    class_service_str.append(
      "\tString update" + class_name + "s (List<" + class_name + "> " + class_name_lower + "s);\n")
    class_service_str.append("\t /** \n \t\0 * 通过主建查询 \n \t\0\0*/\n")
    class_service_str.append(
      "\tList<" + class_name + "> select" + class_name + "(" + class_name + " " + class_name_lower + ");\n")
    class_service_str.append("}\t")
    file_name.write("".join(class_service_str))
    file_name.close()


# serviceImpl
def write_service_impl(class_names, seprarator, path_list):
  for class_name in class_names:
    class_name = class_name[:-5]
    class_service_impl_str = []
    class_name_lower = class_name[0].lower() + class_name[1:]
    file_name = open(path_list[4] + seprarator + class_name + "ServiceImpl.java", "w", encoding="UTF-8")
    class_service_impl_str.append("package " + ".".join(path_list[4].split("""\\""")[4:]) + ";\n\n")
    class_service_impl_str.append("import java.util.List;\n")
    class_service_impl_str.append("import org.springframework.stereotype.Service;\n")
    class_service_impl_str.append("import lombok.extern.slf4j.Slf4j;\n")
    class_service_impl_str.append("import org.springframework.beans.factory.annotation.Autowired;\n")
    class_service_impl_str.append("import " + ".".join(path_list[1].split("""\\""")[4:]) + "." + class_name + ";\n")
    class_service_impl_str.append(
      "import " + ".".join(path_list[2].split("""\\""")[4:]) + "." + class_name + "Dao;\n")
    class_service_impl_str.append(
      "import " + ".".join(path_list[3].split("""\\""")[4:]) + "." + class_name + "Service;\n\n")
    class_service_impl_str.append("@Service(value=\"" + class_name_lower + "Service\")\n")
    class_service_impl_str.append("@Slf4j\n")
    class_service_impl_str.append(
      "public class " + class_name + "ServiceImpl implements " + class_name + "Service{ \n\n")
    class_service_impl_str.append("\t@Autowired\n")
    class_service_impl_str.append("\t" + class_name + "Dao\t" + class_name_lower + "Dao;\n")
    class_service_impl_str.append("")
    class_service_impl_str.append("\t/** \n \t\0* 增加一条 \n \t\0*/\n")
    class_service_impl_str.append("\t@Override\n")
    class_service_impl_str.append(
      "\tpublic String insert" + class_name + " (" + class_name + " " + class_name_lower + "){\n")
    class_service_impl_str.append(
      "\t\t" + class_name_lower + "Dao.insert" + class_name + " (" + class_name_lower + ");\n")
    class_service_impl_str.append("\t\treturn \"\";\n")
    class_service_impl_str.append("\t}\n")
    class_service_impl_str.append("\t/** \n \t\0* 修改一条 \n \t\0*/\n")
    class_service_impl_str.append("\t@Override\n")
    class_service_impl_str.append(
      "\tpublic String update" + class_name + " (" + class_name + " " + class_name_lower + "){\n")
    class_service_impl_str.append(
      "\t\t" + class_name_lower + "Dao.update" + class_name + " (" + class_name_lower + ");\n")
    class_service_impl_str.append("\t\treturn \"\";\n")
    class_service_impl_str.append("\t}\n")
    class_service_impl_str.append("\t/** \n \t\0* 删除一条 \n \t\0*/\n")
    class_service_impl_str.append("\t@Override\n")
    class_service_impl_str.append(
      "\tpublic String delete" + class_name + " (" + class_name + " " + class_name_lower + "){\n")
    class_service_impl_str.append(
      "\t\t" + class_name_lower + "Dao.delete" + class_name + " (" + class_name_lower + ");\n")
    class_service_impl_str.append("\t\treturn \"\";\n")
    class_service_impl_str.append("\t}\n")
    class_service_impl_str.append("\t/** \n \t\0* 增加多条 \n \t\0*/\n")
    class_service_impl_str.append("\t@Override\n")
    class_service_impl_str.append(
      "\tpublic String insert" + class_name + "s (List<" + class_name + "> " + class_name_lower + "s){\n")
    class_service_impl_str.append(
      "\t\t" + class_name_lower + "Dao.insert" + class_name + "s (" + class_name_lower + "s);\n")
    class_service_impl_str.append("\t\treturn \"\";\n")
    class_service_impl_str.append("\t}\n")
    class_service_impl_str.append("\t/** \n \t\0* 修改多条 \n \t\0*/\n")
    class_service_impl_str.append("\t@Override\n")
    class_service_impl_str.append(
      "\tpublic String update" + class_name + "s (List<" + class_name + "> " + class_name_lower + "s){\n")
    class_service_impl_str.append(
      "\t\t" + class_name_lower + "Dao.update" + class_name + "s (" + class_name_lower + "s);\n")
    class_service_impl_str.append("\t\treturn \"\";\n")
    class_service_impl_str.append("\t}\n")
    class_service_impl_str.append("\t/** \n \t\0* 通过主建查询 \n \t\0*/\n")
    class_service_impl_str.append("\t@Override\n")
    class_service_impl_str.append(
      "\tpublic List<" + class_name + "> select" + class_name + "(" + class_name + " " + class_name_lower + "){\n")
    class_service_impl_str.append(
      "\t\treturn " + class_name_lower + "Dao.select" + class_name + " (" + class_name_lower + ");\n")
    class_service_impl_str.append("\t}\n")
    class_service_impl_str.append("}\t")
    file_name.write("".join(class_service_impl_str))
    file_name.close()


# controller
def write_controller(class_names, seprarator, class_strs, path_list):
  for (class_name, class_str) in zip(class_names, class_strs):
    class_name = class_name[:-5]
    class_controller_str = []
    class_name_lower = class_name[0].lower() + class_name[1:]
    file_name = open(path_list[5] + seprarator + class_name + "Controller.java", "w", encoding="UTF-8")
    class_controller_str.append("package " + ".".join(path_list[5].split("""\\""")[4:]) + ";\n\n")
    class_controller_str.append("import java.util.List;\n")
    class_controller_str.append("import org.springframework.beans.factory.annotation.Autowired;\n")
    class_controller_str.append("import lombok.extern.slf4j.Slf4j;\n")
    class_controller_str.append("import org.springframework.web.bind.annotation.RestController;\n")
    class_controller_str.append("import org.springframework.web.bind.annotation.ResponseBody;\n")
    class_controller_str.append("import org.springframework.web.bind.annotation.RequestMapping;\n")
    class_controller_str.append("import org.springframework.web.bind.annotation.RequestBody;\n")
    class_controller_str.append("import org.springframework.web.bind.annotation.PostMapping;\n")
    class_controller_str.append("import " + ".".join(path_list[1].split("""\\""")[4:]) + "." + class_name + ";\n")
    class_controller_str.append(
      "import " + ".".join(path_list[3].split("""\\""")[4:]) + "." + class_name + "Service;\n\n")
    class_controller_str.append(
      "@RestController\n@Slf4j\n@RequestMapping(value = \"/" + class_str.lower() + "\")\n")
    class_controller_str.append("public class " + class_name + "Controller { \n")
    class_controller_str.append("\t@Autowired\n")
    class_controller_str.append("\t" + class_name + "Service " + class_name_lower + "Service;\n")
    class_controller_str.append("\t/** \n \t\0* 增加一条 \n \t\0*/\n")
    class_controller_str.append("\t@PostMapping(value = \"/insert_" + class_str.lower() + "\")\n")
    class_controller_str.append("\t@ResponseBody\n")
    class_controller_str.append(
      "\tpublic String insert" + class_name + " (@RequestBody " + class_name + " " + class_name_lower + "){\n")
    class_controller_str.append(
      "\t\treturn " + class_name_lower + "Service.insert" + class_name + "(" + class_name_lower + ");\n\t}\n")
    class_controller_str.append("\t/** \n \t\0* 修改一条 \n \t\0*/\n")
    class_controller_str.append("\t@PostMapping(value = \"/update_" + class_str.lower() + "\")\n")
    class_controller_str.append("\t@ResponseBody\n")
    class_controller_str.append(
      "\tpublic String update" + class_name + " (@RequestBody " + class_name + " " + class_name_lower + "){\n")
    class_controller_str.append(
      "\t\treturn " + class_name_lower + "Service.update" + class_name + "(" + class_name_lower + ");\n\t}\n")
    class_controller_str.append("\t/** \n \t\0* 删除一条 \n \t\0*/\n")
    class_controller_str.append("\t@PostMapping(value = \"/delete_" + class_str.lower() + "\")\n")
    class_controller_str.append("\t@ResponseBody\n")
    class_controller_str.append(
      "\tpublic String delete" + class_name + " (@RequestBody " + class_name + " " + class_name_lower + "){\n")
    class_controller_str.append(
      "\t\treturn " + class_name_lower + "Service.delete" + class_name + "(" + class_name_lower + ");\n\t}\n")
    class_controller_str.append("\t/** \n \t\0* 增加多条 \n \t\0*/\n")
    class_controller_str.append("\t@PostMapping(value = \"/insert_" + class_str.lower() + "_list\")\n")
    class_controller_str.append("\t@ResponseBody\n")
    class_controller_str.append(
      "\tpublic String insert" + class_name + "s (@RequestBody List<" + class_name + "> " + class_name_lower + "s){\n")
    class_controller_str.append(
      "\t\treturn " + class_name_lower + "Service.insert" + class_name + "s(" + class_name_lower + "s);\n\t}\n")
    class_controller_str.append("\t/** \n \t\0* 修改多条 \n \t\0*/\n")
    class_controller_str.append("\t@PostMapping(value = \"/update_" + class_str.lower() + "_list\")\n")
    class_controller_str.append("\t@ResponseBody\n")
    class_controller_str.append(
      "\tpublic String update" + class_name + "s (@RequestBody List<" + class_name + "> " + class_name_lower + "s){\n")
    class_controller_str.append(
      "\t\treturn " + class_name_lower + "Service.update" + class_name + "s(" + class_name_lower + "s);\n\t}\n")
    class_controller_str.append("\t/** \n \t\0* 通过主建查询 \n \t\0*/\n")
    class_controller_str.append("\t@PostMapping(value = \"/select_" + class_str.lower() + "_list\")\n")
    class_controller_str.append("\t@ResponseBody\n")
    class_controller_str.append(
      "\tpublic List<" + class_name + "> select" + class_name + "(@RequestBody " + class_name + " " + class_name_lower + "){\n")
    class_controller_str.append(
      "\t\treturn " + class_name_lower + "Service.select" + class_name + "(" + class_name_lower + ");\n\t}\n")
    class_controller_str.append("}\t")
    file_name.write("".join(class_controller_str))
    file_name.close()


# application 文件
def write_application(path, seprarator):
  mapper_file_yml = open(path + seprarator + "applicatoin.yml", 'w', encoding="utf-8")
  mapper_file_properties = open(path + seprarator + "applicatoin.properties", 'w', encoding="utf-8")
  mapper_file_yml.close()
  mapper_file_properties.close()


# mapper
def write_mapper(sheet, seprarator, path_list):
  global rows_num
  global rows_num_start
  # mapper开头拼接
  # = []
  while rows_num < sheet.nrows:
    if sheet.row_values(rows_num)[1] is '':
      rows_num += 1
      continue
    cl_class_old = str(sheet.row_values(rows_num)[1])
    cl_class_new = string.capwords(cl_class_old.replace('_', ' '))
    class_name_str = cl_class_new.replace(' ', '')
    table_name_str = str(sheet.row_values(rows_num)[1])
    mapper_file = open(path_list[6] + seprarator + class_name_str + "Mapper.xml", 'w', encoding="utf-8")
    param_map = {}

    map_toung = []
    map_toung.append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
    map_toung.append(
      "<!DOCTYPE mapper PUBLIC \"-//mybatis.org//DTD Mapper 3.0//EN\" \"http://mybatis.org/dtd/mybatis-3-mapper.dtd\" >\n")
    map_toung.append(
      "<mapper namespace=\"" + ".".join(path_list[2].split("""\\""")[4:]) + "." + class_name_str + "Dao\">\n")
    # insert 拼接
    insert_sql_one = []
    insert_sql_list = []
    insert_sql_same = []
    # select拼接
    select_sql = []
    # update拼接
    update_sql_one = []
    update_sql_list = []
    update_sql_same = []
    select_sql.append("<select id=\"select" + class_name_str + "\" parameterType=\"" + ".".join(
      path_list[1].split("""\\""")[4:]) + "." + class_name_str +
                      "\"\n\tresultType=\"" + ".".join(
      path_list[1].split("""\\""")[4:]) + "." + class_name_str + "\">\n\tSELECT \n")
    insert_sql_list.append("<insert id=\"insert" + class_name_str + "\" parameterType=\"java.util.List\">\n")
    insert_sql_one.append("<insert id=\"insert" + class_name_str + "s\" parameterType=\"" + ".".join(
      path_list[1].split("""\\""")[4:]) + "." + class_name_str + "\">\n")
    insert_sql_toung = []
    insert_sql_toung.append("\tINSERT INTO " + table_name_str + " (\n")
    insert_sql_middle = []
    insert_sql_end = []
    for i in range(rows_num, sheet.nrows):
      rows_num += 1
      if sheet.row_values(i)[1] is '':
        # property_str = []
        break
      if sheet.row_values(i)[3] is '':
        continue
      cl_db = str(sheet.row_values(i)[1])
      if i == sheet.nrows - 1 or sheet.row_values(i + 1)[1] is '':
        insert_sql_toung.append("\t\t" + cl_db + "\n")
        select_sql.append("\t\t" + cl_db + "\n")
      else:
        insert_sql_toung.append("\t\t" + cl_db + ",\n")
        select_sql.append("\t\t" + cl_db + ",\n")

      cl_class_old = str(sheet.row_values(i)[1])
      cl_class_new = string.capwords(cl_class_old.replace('_', ' '))
      cl_str_property = cl_class_new.replace(' ', '')
      cl_str = cl_str_property[0].lower() + cl_str_property[1:]
      cl_type = ""
      if i == 0:
        new_str = cl_str_property
        continue
      # insert批量 拼接
      if "VARCHAR" in str(sheet.row_values(i)[3]):
        cl_type = "VARCHAR"
      elif "DATE" in str(sheet.row_values(i)[3]):
        cl_type = "DATE"
      elif "NUMBER" in str(sheet.row_values(i)[3]) or "DECIMAL" in str(sheet.row_values(i)[3]):
        cl_type = "DOUBLE"
      if i == sheet.nrows - 1 or sheet.row_values(i + 1)[1] is '':
        if "LAST_UPDATE_TIME" == str(sheet.row_values(i)[1]) or "CREATE_DATE" == str(sheet.row_values(i)[1]):
          insert_sql_middle.append("\t\tSYSDATE\n")
        else:
          insert_sql_middle.append("\t\t#{item." + cl_str + ", jdbcType=" + cl_type + "}\n")
      else:
        if "LAST_UPDATE_TIME" == str(sheet.row_values(i)[1]) or "CREATE_DATE" == str(sheet.row_values(i)[1]):
          insert_sql_middle.append("\t\tSYSDATE,\n")
        elif str(sheet.row_values(i)[5]) == "N":
          insert_sql_middle.append("\t\t#{item." + cl_str + "},\n")
        else:
          insert_sql_middle.append("\t\t#{item." + cl_str + ", jdbcType=" + cl_type + "},\n")

      # update批量拼接
      if not str(sheet.row_values(i)[4]) == "":
        param_map[cl_class_old] = cl_str
        continue
      if i == sheet.nrows - 1 or sheet.row_values(i + 1)[1] is '':
        if "VERSION" == cl_class_old:
          update_sql_same.append('\t\t' + cl_class_old + '=' + cl_class_old + '+1\n')
        elif "LAST_UPDATE_TIME" == cl_class_old:
          update_sql_same.append('\t\tSYSDATE\n')
        else:
          update_sql_same.append("\t\t<if test=\"item." + cl_str + "!=null and item." + cl_str + " !=''\">\n")
          update_sql_same.append("\t\t\t" + cl_class_old + "=" + "#{item." + cl_str + "}\n")
          update_sql_same.append("\t\t</if>\n")
      else:
        if "VERSION" == cl_class_old:
          update_sql_same.append('\t\t' + cl_class_old + '=' + cl_class_old + '+1,\n')
        elif "LAST_UPDATE_TIME" == cl_class_old:
          update_sql_same.append('\t\tSYSDATE,\n')
        else:
          update_sql_same.append("\t\t<if test=\"item." + cl_str + "!=null and item." + cl_str + " !=''\">\n")
          update_sql_same.append("\t\t\t" + cl_class_old + "=" + "#{item." + cl_str + "},\n")
          update_sql_same.append("\t\t</if>\n")

    select_sql.append("\tFROM " + table_name_str + " WHERE \n")
    insert_sql_toung.append("\t) ( \n")
    insert_sql_one.append("".join(insert_sql_toung))
    insert_sql_list.append("".join(insert_sql_toung))
    insert_sql_list.append(
      "\t <foreach item=\"item\" collection=\"list\" index=\"index\" separator=\"union all \">\n")
    insert_sql_list.append("\tSELECT \n")
    insert_sql_one.append("".join(insert_sql_middle))
    insert_sql_list.append("".join(insert_sql_middle))
    insert_sql_list.append("\tfrom dual \n")
    insert_sql_list.append("\t</foreach>\n")
    insert_sql_list.append("\t)\n")
    insert_sql_list.append("</insert>\n")

    insert_sql_one.append("\t)\n")
    insert_sql_one.append("</insert>\n")

    update_sql_list.append("<update id=\"update" + class_name_str + "\" parameterType=\"java.util.List\">\n")
    update_sql_list.append("\tBEGIN\n")
    update_sql_list.append("\t<foreach collection=\"list\" separator=\";\" item=\"item\" index=\"index\">\n")
    update_sql_list.append("\t\tUPDATE " + table_name_str + " SET \n")
    update_sql_one.append("<update id=\"update" + class_name_str + "s\" parameterType=\"" + ".".join(
      path_list[1].split("""\\""")[4:]) + "." + class_name_str + "\">\n")
    update_sql_one.append("\t\tUPDATE " + table_name_str + " SET \n")

    update_sql_same.append("\t\tWHERE\n")
    i = 0
    # delete拼接
    delete_sql_list = []
    delete_sql_list.append("<delete id=\"delete" + cl_str + "\" parameterType=\" " + ".".join(
      path_list[1].split("""\\""")[4:]) + "." + class_name_str + "\">\n")
    delete_sql_list.append("\tdelete from " + table_name_str + " where \n")
    delete_sql_list.append("\t<foreach collection=\"list\" separator=\"or\" item=\"item\" index=\"index\">\n")
    for key, value in param_map.items():
      i += 1
      if i == len(param_map):
        update_sql_same.append("\t\t" + key + "=#{item." + value + "}\n")
        delete_sql_list.append("\t\t" + key + "=#{item." + value + "}\n")
        select_sql.append("\t\t" + key + "=#{item." + value + "}\n")
        break
      update_sql_same.append("\t\t" + key + "=#{item." + value + "} and \n")
      delete_sql_list.append("\t\t(" + key + "=#{item." + value + "} and \n")
      select_sql.append("\t\t" + key + "=#{" + value + "} and \n")

    update_sql_list.append("".join(update_sql_same))
    update_sql_list.append("\t</foreach>\n")
    update_sql_list.append("\t;END;\n")
    update_sql_list.append("</update>\n")
    update_sql_one.append("".join(update_sql_same))
    update_sql_one.append("</update>\n")
    delete_sql_list.append("\t</foreach>\n</delete>\n")
    select_sql.append("</select>")
    mapper_str = map_toung + insert_sql_one + insert_sql_list + update_sql_one + update_sql_list + delete_sql_list + select_sql
    mapper_str.append("\n</mapper>")
    mapper_file.write("".join(mapper_str))
    mapper_file.close()
  rows_num = rows_num_start


# test serviceImpl
def write_test_service_impl(class_names, seprarator, path_list):
  for class_name in class_names:
    class_name = class_name[:-5]
    class_test_service_impl_str = []
    class_name_lower = class_name[0].lower() + class_name[1:]
    file_name = open(path_list[-5] + seprarator + class_name + "ServiceImplTest.java", "w", encoding="UTF-8")
    class_test_service_impl_str.append("package " + ".".join(path_list[-5].split("""\\""")[4:]) + ";\n\n")
    class_test_service_impl_str.append("import java.util.List;\n")
    class_test_service_impl_str.append("import org.springframework.beans.factory.annotation.Autowired;\n")
    class_test_service_impl_str.append(
      "import " + ".".join(path_list[1].split("""\\""")[4:]) + "." + class_name + ";\n")
    class_test_service_impl_str.append(
      "import " + ".".join(path_list[2].split("""\\""")[4:]) + "." + class_name + "Dao;\n")
    class_test_service_impl_str.append(
      "import " + ".".join(path_list[3].split("""\\""")[4:]) + "." + class_name + "Service;\n\n")
    class_test_service_impl_str.append("import org.junit.runner.RunWith;\n")
    class_test_service_impl_str.append("import org.junit.Test;\n")
    class_test_service_impl_str.append("import org.junit.Before;\n")
    class_test_service_impl_str.append("import org.springframework.boot.test.context.SpringBootTest;\n")
    class_test_service_impl_str.append("import org.springframework.test.context.junit4.SpringRunner;\n\n")
    class_test_service_impl_str.append("@RunWith(SpringRunner.class)\n")
    class_test_service_impl_str.append("@SpringBootTest\n")
    class_test_service_impl_str.append("public class " + class_name + "ServiceImplTest { \n\n")
    class_test_service_impl_str.append("\t@Autowired\n")
    class_test_service_impl_str.append("\t" + class_name + "Dao\t" + class_name_lower + "Dao;\n")
    class_test_service_impl_str.append("\t@Autowired\n")
    class_test_service_impl_str.append("\t" + class_name + "Service\t" + class_name_lower + "Service;\n")
    class_test_service_impl_str.append("\t@Before\n")
    class_test_service_impl_str.append("\tpublic void setUp() {\n\n")
    class_test_service_impl_str.append("\t}\n")
    class_test_service_impl_str.append("\t@Test\n")
    class_test_service_impl_str.append("\tpublic void insert" + class_name + " (){\n\n")
    class_test_service_impl_str.append("\t}\n")
    class_test_service_impl_str.append("\t@Test\n")
    class_test_service_impl_str.append("\tpublic void update" + class_name + " (){\n\n")
    class_test_service_impl_str.append("\t}\n")
    class_test_service_impl_str.append("\t@Test\n")
    class_test_service_impl_str.append("\tpublic void delete" + class_name + " (){\n\n")
    class_test_service_impl_str.append("\t}\n")
    class_test_service_impl_str.append("\t@Test\n")
    class_test_service_impl_str.append("\tpublic void insert" + class_name + "List (){\n\n")
    class_test_service_impl_str.append("\t}\n")
    class_test_service_impl_str.append("\t@Test\n")
    class_test_service_impl_str.append("\tpublic void update" + class_name + "List (){\n\n")
    class_test_service_impl_str.append("\t}\n")
    class_test_service_impl_str.append("\t@Test\n")
    class_test_service_impl_str.append("\tpublic void select" + class_name + " (){\n\n")
    class_test_service_impl_str.append("\t}\n")
    class_test_service_impl_str.append("}")
    file_name.write("".join(class_test_service_impl_str))
    file_name.close()


# test controller
def write_test_controller(class_names, seprarator, path_list):
  for class_name in class_names:
    class_name = class_name[:-5]
    class_test_controller_str = []
    class_name_lower = class_name[0].lower() + class_name[1:]
    file_name = open(path_list[-4] + seprarator + class_name + "ControllerTest.java", "w", encoding="UTF-8")
    class_test_controller_str.append("package " + ".".join(path_list[-4].split("""\\""")[4:]) + ";\n\n")
    class_test_controller_str.append("import java.util.List;\n")
    class_test_controller_str.append("import org.springframework.beans.factory.annotation.Autowired;\n")
    class_test_controller_str.append(
      "import " + ".".join(path_list[1].split("""\\""")[4:]) + "." + class_name + ";\n")
    class_test_controller_str.append(
      "import " + ".".join(path_list[3].split("""\\""")[4:]) + "." + class_name + "Service;\n\n")
    class_test_controller_str.append("import org.junit.runner.RunWith;\n")
    class_test_controller_str.append("import org.junit.Test;\n")
    class_test_controller_str.append("import org.junit.Before;\n")
    class_test_controller_str.append("import org.springframework.boot.test.context.SpringBootTest;\n")
    class_test_controller_str.append("import org.springframework.test.context.junit4.SpringRunner;\n\n")
    class_test_controller_str.append("@RunWith(SpringRunner.class)\n")
    class_test_controller_str.append("@SpringBootTest\n")
    class_test_controller_str.append("public class " + class_name + "ControllerTest { \n\n")
    class_test_controller_str.append("\t@Autowired\n")
    class_test_controller_str.append("\t" + class_name + "Service\t" + class_name_lower + "Service;\n")
    class_test_controller_str.append("\t@Before\n")
    class_test_controller_str.append("\tpublic void setUp() {\n\n")
    class_test_controller_str.append("\t}\n")
    class_test_controller_str.append("\t@Test\n")
    class_test_controller_str.append("\tpublic void insert" + class_name + " (){\n\n")
    class_test_controller_str.append("\t}\n")
    class_test_controller_str.append("\t@Test\n")
    class_test_controller_str.append("\tpublic void update" + class_name + " (){\n\n")
    class_test_controller_str.append("\t}\n")
    class_test_controller_str.append("\t@Test\n")
    class_test_controller_str.append("\tpublic void delete" + class_name + " (){\n\n")
    class_test_controller_str.append("\t}\n")
    class_test_controller_str.append("\t@Test\n")
    class_test_controller_str.append("\tpublic void insert" + class_name + "List (){\n\n")
    class_test_controller_str.append("\t}\n")
    class_test_controller_str.append("\t@Test\n")
    class_test_controller_str.append("\tpublic void update" + class_name + "List (){\n\n")
    class_test_controller_str.append("\t}\n")
    class_test_controller_str.append("\t@Test\n")
    class_test_controller_str.append("\tpublic void select" + class_name + " (){\n\n")
    class_test_controller_str.append("\t}\n")
    class_test_controller_str.append("}")
    file_name.write("".join(class_test_controller_str))
    file_name.close()
  pass


# pom文件
def write_pom(sheet, file_path):
  # pom 内容
  pom = []
  pom.append("""<project xmlns="http://maven.apache.org/POM/4.0.0" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
       xsi:schemaLocation="http://maven.apache.org/POM/4.0.0 http://maven.apache.org/xsd/maven-4.0.0.xsd">""")
  wfile_pom = open(file_path, "w", encoding="utf-8")
  pom.append("\n\t<modelVersion>4.0.0</modelVersion>")
  # pom.append("\n\t<parent>")
  # for i in range(4, 7):
  #   pom.append(
  #     "\n\t\t<" + sheet.row_values(i)[2] + ">" + sheet.row_values(i)[3] + "</" + sheet.row_values(i)[2] + ">")
  # pom.append("\n\t</parent>")
  # pom.append("\n\t<" + sheet.row_values(7)[1] + ">" + sheet.row_values(7)[3] + "</" + sheet.row_values(7)[1] + ">")
  pom.append("\n</project>")
  wfile_pom.write("".join(pom))
  wfile_pom.close()


# 启动文件
def write_starter(seprarator, path):
  starter_name = path.split("""\\""")[0]
  starter_name = string.capwords(starter_name.replace('-', ' ')).replace(' ', '')
  starter_file = open(path + seprarator + starter_name + 'ApplicationStarter.java', "w", encoding="utf-8")
  starter_str = []
  starter_str.append("package " + ".".join(path.split("""\\""")[4:]) + ";\n\n")
  starter_str.append("import org.springframework.boot.SpringApplication;\n")
  starter_str.append("import org.springframework.boot.autoconfigure.SpringBootApplication;\n")
  starter_str.append("import org.mybatis.spring.annotation.MapperScan;\n\n")
  starter_str.append("@SpringBootApplication\n")
  starter_str.append("@MapperScan(basePackages={\"" + ".".join(path.split("""\\""")[4:]) + ".**.dao\"})\n")
  starter_str.append("public class " + starter_name + "ApplicationStarter {\n")
  starter_str.append("\tpublic static void main(String[] args) {\n")
  starter_str.append("\t\tSpringApplication.run(" + starter_name + "ApplicationStarter.class, args);\n")
  starter_str.append("\t}\n}")
  starter_file.write("".join(starter_str))
  starter_file.close()


# 创建目录
def make_dirs(sheet):
  # 需要创建文件的文件夹
  project_must_all_path_list = []
  # 项目文件夹
  project_path = str(sheet.row_values(0)[1])
  project_must_all_path_list.append(project_path)
  # java根目录
  root_java_path = """\\src\\main\\java\\"""
  # resource根目录
  root_resources_path = """\\src\\main\\resources"""
  # test根目录
  root_test_path = """\\src\\test\\java\\"""
  # 基础目录
  project_base_path = str(sheet.row_values(1)[1])
  # 必须目录
  project_must_path = str(sheet.row_values(2)[1])
  # resouces必须目录
  resouces_must_path = str(sheet.row_values(3)[1])
  # 目录分割符
  seprarator = """\\"""
  # 基本目录
  base_java_path = project_path + root_java_path + project_base_path.replace(".", seprarator)
  base_resources_path = project_path + root_resources_path
  base_test_path = project_path + root_test_path + project_base_path.replace(".", seprarator)
  # 创建项目目录
  if not os.path.exists(project_path):
    os.makedirs(project_path)
  # 创建java目录
  if not os.path.exists(base_java_path):
    os.makedirs(base_java_path)
  project_must_path_list = project_must_path.split(",")
  resouces_must_path_list = resouces_must_path.split(",")
  # 创建resouces目录
  if not os.path.exists(base_resources_path):
    os.makedirs(base_resources_path)
  # 创建test目录
  if not os.path.exists(base_test_path):
    os.makedirs(base_test_path)
  # 创建项目java基础目录
  for path in project_must_path_list:
    path = path.replace(".", seprarator)
    if not os.path.exists(base_java_path + seprarator + path):
      project_must_all_path_list.append(base_java_path + seprarator + path)
      os.makedirs(base_java_path + seprarator + path)
  # 创建resource基础目录
  for path in resouces_must_path_list:
    if not os.path.exists(base_resources_path + seprarator + path):
      project_must_all_path_list.append(base_resources_path + seprarator + path)
      os.makedirs(base_resources_path + seprarator + path)
      # 创建项目java基础目录
  # 创建test基础目录
  for path in project_must_path_list:
    path = path.replace(".", seprarator)
    if not os.path.exists(base_test_path + seprarator + path):
      project_must_all_path_list.append(base_test_path + seprarator + path)
      os.makedirs(base_test_path + seprarator + path)
  project_must_all_path_list.append(base_resources_path)
  project_must_all_path_list.append(base_java_path)
  return project_must_all_path_list


if __name__ == '__main__':
  main()
