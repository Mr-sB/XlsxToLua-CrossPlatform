using XlsxToLua.Json;
using System;
using System.Collections.Generic;
using System.Text;

namespace XlsxToLua
{
// Extended part for special exporting requirement
    public partial class TableExportToJsonHelper
    {
        /// <summary>
        /// 按配置的特殊索引导出方式输出json文件
        /// </summary>
        public static bool SpecialExportTableToJson(TableInfo tableInfo, string exportRule, out string errorString)
        {
            exportRule = exportRule.Trim();
            // 解析按这种方式导出后的json文件名
            int colonIndex = exportRule.IndexOf(':');
            if (colonIndex == -1)
            {
                errorString = string.Format("导出配置\"{0}\"定义错误，必须在开头声明导出json文件名\n", exportRule);
                return false;
            }

            string fileName = exportRule.Substring(0, colonIndex).Trim();
            // 判断是否在最后的花括号内声明table value中包含的字段
            int leftBraceIndex = exportRule.LastIndexOf('{');
            int rightBraceIndex = exportRule.LastIndexOf('}');
            // 解析依次作为索引的字段名
            string indexFieldNameString = null;
            // 注意分析花括号时要考虑到未声明table value中的字段而在某索引字段完整性检查规则中用花括号声明了有效值的情况
            if (exportRule.EndsWith("}") && leftBraceIndex != -1)
                indexFieldNameString = exportRule.Substring(colonIndex + 1, leftBraceIndex - colonIndex - 1);
            else
                indexFieldNameString = exportRule.Substring(colonIndex + 1, exportRule.Length - colonIndex - 1);

            string[] indexFieldDefine = indexFieldNameString.Split(new char[] {'-'}, System.StringSplitOptions.RemoveEmptyEntries);
            // 用于索引的字段列表
            List<FieldInfo> indexField = new List<FieldInfo>();
            // 索引字段对应的完整性检查规则
            List<string> integrityCheckRules = new List<string>();
            if (indexFieldDefine.Length < 1)
            {
                errorString = string.Format(
                    "导出配置\"{0}\"定义错误，用于索引的字段不能为空，请按fileName:indexFieldName1-indexFieldName2{otherFieldName1,otherFieldName2}的格式配置\n", exportRule);
                return false;
            }

            // 检查字段是否存在且为int、float、string或lang型
            foreach (string fieldDefine in indexFieldDefine)
            {
                string fieldName = null;
                // 判断是否在字段名后用小括号声明了该字段的完整性检查规则
                int leftBracketIndex = fieldDefine.IndexOf('(');
                int rightBracketIndex = fieldDefine.IndexOf(')');
                if (leftBracketIndex > 0 && rightBracketIndex > leftBracketIndex)
                {

                    fieldName = fieldDefine.Substring(0, leftBracketIndex);
                    string integrityCheckRule = fieldDefine.Substring(leftBracketIndex + 1, rightBracketIndex - leftBracketIndex - 1).Trim();
                    if (string.IsNullOrEmpty(integrityCheckRule))
                    {
                        errorString = string.Format("导出配置\"{0}\"定义错误，用于索引的字段\"{1}\"若要声明完整性检查规则就必须在括号中填写否则不要加括号\n", exportRule, fieldName);
                        return false;
                    }

                    integrityCheckRules.Add(integrityCheckRule);
                }
                else
                {
                    fieldName = fieldDefine.Trim();
                    integrityCheckRules.Add(null);
                }

                FieldInfo fieldInfo = tableInfo.GetFieldInfoByFieldName(fieldName);
                if (fieldInfo == null)
                {
                    errorString = string.Format("导出配置\"{0}\"定义错误，声明的索引字段\"{1}\"不存在\n", exportRule, fieldName);
                    return false;
                }

                if (fieldInfo.DataType != DataType.Int && fieldInfo.DataType != DataType.Long && fieldInfo.DataType != DataType.Float &&
                    fieldInfo.DataType != DataType.String && fieldInfo.DataType != DataType.Lang)
                {
                    errorString = string.Format("导出配置\"{0}\"定义错误，声明的索引字段\"{1}\"为{2}型，但只允许为int、long、float、string或lang型\n", exportRule, fieldName,
                        fieldInfo.DataType);
                    return false;
                }

                // 对索引字段进行非空检查
                if (fieldInfo.DataType == DataType.String)
                {
                    FieldCheckRule stringNotEmptyCheckRule = new FieldCheckRule();
                    stringNotEmptyCheckRule.CheckType = TableCheckType.NotEmpty;
                    stringNotEmptyCheckRule.CheckRuleString = "notEmpty[trim]";
                    TableCheckHelper.CheckNotEmpty(fieldInfo, stringNotEmptyCheckRule, out errorString);
                    if (errorString != null)
                    {
                        errorString = string.Format("按配置\"{0}\"进行自定义导出错误，string型索引字段\"{1}\"中存在以下空值，而作为索引的key不允许为空\n{2}\n", exportRule, fieldName,
                            errorString);
                        return false;
                    }
                }
                else if (fieldInfo.DataType == DataType.Lang)
                {
                    FieldCheckRule langNotEmptyCheckRule = new FieldCheckRule();
                    langNotEmptyCheckRule.CheckType = TableCheckType.NotEmpty;
                    langNotEmptyCheckRule.CheckRuleString = "notEmpty[key|value]";
                    TableCheckHelper.CheckNotEmpty(fieldInfo, langNotEmptyCheckRule, out errorString);
                    if (errorString != null)
                    {
                        errorString = string.Format("按配置\"{0}\"进行自定义导出错误，lang型索引字段\"{1}\"中存在以下空值，而作为索引的key不允许为空\n{2}\n", exportRule, fieldName,
                            errorString);
                        return false;
                    }
                }
                else if (AppValues.IsAllowedNullNumber == true)
                {
                    FieldCheckRule numberNotEmptyCheckRule = new FieldCheckRule();
                    numberNotEmptyCheckRule.CheckType = TableCheckType.NotEmpty;
                    numberNotEmptyCheckRule.CheckRuleString = "notEmpty";
                    TableCheckHelper.CheckNotEmpty(fieldInfo, numberNotEmptyCheckRule, out errorString);
                    if (errorString != null)
                    {
                        errorString = string.Format("按配置\"{0}\"进行自定义导出错误，{1}型索引字段\"{2}\"中存在以下空值，而作为索引的key不允许为空\n{3}\n", exportRule,
                            fieldInfo.DataType.ToString(), fieldName, errorString);
                        return false;
                    }
                }

                indexField.Add(fieldInfo);
            }

            // 解析table value中要输出的字段名
            List<FieldInfo> tableValueField = new List<FieldInfo>();
            // 如果在花括号内配置了table value中要输出的字段名
            if (exportRule.EndsWith("}") && leftBraceIndex != -1 && leftBraceIndex < rightBraceIndex)
            {
                string tableValueFieldName = exportRule.Substring(leftBraceIndex + 1, rightBraceIndex - leftBraceIndex - 1);
                string[] fieldNames = tableValueFieldName.Split(new char[] {','}, System.StringSplitOptions.RemoveEmptyEntries);
                if (fieldNames.Length < 1)
                {
                    errorString = string.Format(
                        "导出配置\"{0}\"定义错误，花括号中声明的table value中的字段不能为空，请按fileName:indexFieldName1-indexFieldName2{otherFieldName1,otherFieldName2}的格式配置\n",
                        exportRule);
                    return false;
                }

                // 检查字段是否存在
                foreach (string fieldName in fieldNames)
                {
                    FieldInfo fieldInfo = tableInfo.GetFieldInfoByFieldName(fieldName);
                    if (fieldInfo == null)
                    {
                        errorString = string.Format("导出配置\"{0}\"定义错误，声明的table value中的字段\"{1}\"不存在\n", exportRule, fieldName);
                        return false;
                    }

                    if (tableValueField.Contains(fieldInfo))
                        Utils.LogWarning(string.Format("警告：导出配置\"{0}\"定义中，声明的table value中的字段存在重复，字段名为{1}（列号{2}），本工具只生成一次，请修正错误\n", exportRule,
                            fieldName, Utils.GetExcelColumnName(fieldInfo.ColumnSeq + 1)));
                    else
                        tableValueField.Add(fieldInfo);
                }
            }
            else if (exportRule.EndsWith("}") && leftBraceIndex == -1)
            {
                errorString = string.Format("导出配置\"{0}\"定义错误，声明的table value中花括号不匹配\n", exportRule);
                return false;
            }
            // 如果未在花括号内声明，则默认将索引字段之外的所有字段进行填充
            else
            {
                List<string> indexFieldNameList = new List<string>(indexFieldDefine);
                foreach (FieldInfo fieldInfo in tableInfo.GetAllClientFieldInfo())
                {
                    if (!indexFieldNameList.Contains(fieldInfo.FieldName))
                        tableValueField.Add(fieldInfo);
                }
            }

            // 解析完依次作为索引的字段以及table value中包含的字段后，按索引要求组成相应的嵌套数据结构
            Dictionary<object, object> data = new Dictionary<object, object>();
            int rowCount = tableInfo.GetKeyColumnFieldInfo().Data.Count;
            for (int i = 0; i < rowCount; ++i)
            {
                Dictionary<object, object> temp = data;
                // 生成除最内层的数据结构
                for (int j = 0; j < indexField.Count - 1; ++j)
                {
                    FieldInfo oneIndexField = indexField[j];
                    var tempData = oneIndexField.Data[i];
                    if (!temp.ContainsKey(tempData))
                        temp.Add(tempData, new Dictionary<object, object>());

                    temp = (Dictionary<object, object>) temp[tempData];
                }

                // 最内层的value存数据的int型行号（从0开始计）
                FieldInfo lastIndexField = indexField[indexField.Count - 1];
                var lastIndexFieldData = lastIndexField.Data[i];
                if (!temp.ContainsKey(lastIndexFieldData))
                    temp.Add(lastIndexFieldData, i);
                else
                {
                    errorString = string.Format("错误：对表格{0}按\"{1}\"规则进行特殊索引导出时发现第{2}行与第{3}行在各个索引字段的值完全相同，导出被迫停止，请修正错误后重试\n", tableInfo.TableName,
                        exportRule, i + AppValues.DATA_FIELD_DATA_START_INDEX + 1, temp[lastIndexFieldData]);
                    Utils.LogErrorAndExit(errorString);
                    return false;
                }
            }

            // 进行数据完整性检查
            if (AppValues.IsNeedCheck == true)
            {
                TableCheckHelper.CheckTableIntegrity(indexField, data, integrityCheckRules, out errorString);
                if (errorString != null)
                {
                    errorString = string.Format("错误：对表格{0}按\"{1}\"规则进行特殊索引导时未通过数据完整性检查，导出被迫停止，请修正错误后重试：\n{2}\n", tableInfo.TableName, exportRule,
                        errorString);
                    return false;
                }
            }

            // 生成导出的文件内容
            StringBuilder content = new StringBuilder();

            // 生成json字符串开头
            content.Append("{");

            // 逐层按嵌套结构输出数据
            _GetIndexFieldData(content, data, tableValueField, out errorString);
            if (errorString != null)
            {
                errorString = string.Format("错误：对表格{0}按\"{1}\"规则进行特殊索引导出时发现以下错误，导出被迫停止，请修正错误后重试：\n{2}\n", tableInfo.TableName, exportRule,
                    errorString);
                return false;
            }

            // 去掉最后一行后多余的英文逗号，此处要特殊处理当表格中没有任何数据行时的情况
            if (content.Length > 1)
                content.Remove(content.Length - 1, 1);
            // 生成数据内容结尾
            content.Append("}");

            string exportString = content.ToString();
            // 如果声明了要整理为带缩进格式的形式
            if (AppValues.ExportJsonIsFormat == true)
                exportString = _FormatJson(exportString);

            // 保存为json文件
            if (Utils.SaveJsonFile(tableInfo.TableName, fileName, exportString) == true)
            {
                errorString = null;
                return true;
            }
            else
            {
                errorString = "保存为json文件失败\n";
                return false;
            }
        }

        /// <summary>
        /// 按指定索引方式导出数据时,通过此函数递归生成层次结构,当递归到最内层时输出指定table value中的数据
        /// </summary>
        static void _GetIndexFieldData(StringBuilder content, Dictionary<object, object> parentDict, List<FieldInfo> tableValueField,
            out string errorString)
        {
            foreach (var key in parentDict.Keys)
            {
                // 生成key
                if (key.GetType() == typeof(int) || key.GetType() == typeof(float))
                {
                    content.Append("\"").Append(key).Append("\"");
                }
                else if (key.GetType() == typeof(string))
                {
                    //// 检查作为key值的变量名是否合法
                    //TableCheckHelper.CheckFieldName(key.ToString(), out errorString);
                    //if (errorString != null)
                    //{
                    //    errorString = string.Format("作为第{0}层索引的key值不是合法的变量名，你填写的为\"{1}\"", currentLevel - 1, key.ToString());
                    //    return;
                    //}
                    //content.Append(key);

                    content.Append("\"").Append(key).Append("\"");
                }
                else
                {
                    errorString = string.Format("SpecialExportTableToLua中出现非法类型的索引列类型{0}", key.GetType());
                    Utils.LogErrorAndExit(errorString);
                    return;
                }

                content.Append(":{");

                string notUsed;
                // 如果已是最内层，输出指定table value中的数据
                if (parentDict[key].GetType() == typeof(int))
                {
                    foreach (FieldInfo fieldInfo in tableValueField)
                    {
                        int rowIndex = (int) parentDict[key];
                        string oneTableValueFieldData = _GetOneField(fieldInfo, rowIndex, out errorString, out notUsed);
                        if (errorString != null)
                        {
                            errorString = string.Format("第{0}行的字段\"{1}\"（列号：{2}）导出数据错误：{3}", rowIndex + AppValues.DATA_FIELD_DATA_START_INDEX + 1,
                                fieldInfo.FieldName, Utils.GetExcelColumnName(fieldInfo.ColumnSeq + 1), errorString);
                            return;
                        }
                        else
                            content.Append(oneTableValueFieldData);
                    }
                }
                // 否则继续递归生成索引key
                else
                {
                    _GetIndexFieldData(content, (Dictionary<object, object>) (parentDict[key]), tableValueField, out errorString);
                    if (errorString != null)
                        return;
                }

                // 去掉本行最后一个字段后多余的英文逗号，json语法不像lua那样最后一个字段后的逗号可有可无
                content.Remove(content.Length - 1, 1);
                // 生成一行数据json object的结尾
                content.Append("}");
                // 每行的json object后加英文逗号
                content.Append(",");
            }

            errorString = null;
        }
    }
}
