using CARETTA.COM.Infrastructure.Logger;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Reflection;

namespace FPO.COM.Helper
{
    /// <summary>
    /// custom excel data reader to list
    /// matches excel columns with model's ExcelColumn attribute
    /// </summary>
    public class ExcelHelper
    {
        private OleDbConnection Xlsxconnection { get; set; }
        private OleDbCommand Command { get; set; }
        private OleDbDataReader Reader { get; set; }

        public List<T> ReadExcelFromFirstSheet<T>(string source, bool deleteAfterRead = true)
        {

            Log4NetLogger logger = new Log4NetLogger();
            List<T> returnList = new List<T>();
            try
            {
                using (Xlsxconnection = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + source + "; Extended Properties='Excel 12.0 Xml;HDR=YES'"))
                {
                    Xlsxconnection.Open();
                    var dtSchema = Xlsxconnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                    string firstSheetName = dtSchema.Rows[0].Field<string>("TABLE_NAME");
                    using (Command = new OleDbCommand("SELECT * FROM [" + firstSheetName + "]", Xlsxconnection))
                    {
                        using (Reader = Command.ExecuteReader())
                        {
                            var properties = typeof(T).GetProperties();
                            while (Reader.Read())
                            {
                                var instance = Activator.CreateInstance(typeof(T));
                                Dictionary<string, string> propAttributes = GetPropertyAttributes<T>();
                                foreach (var property in properties)
                                {
                                    PropertyInfo instanceProperty = instance.GetType().GetProperty(property.Name);
                                    string attribute = string.Empty;
                                    propAttributes.TryGetValue(property.Name, out attribute);
                                    if (!string.IsNullOrEmpty(attribute))
                                    {
                                        instanceProperty.SetValue(instance, GetConvertedValue(Reader[attribute].ToString(), instanceProperty.PropertyType), null);
                                    }
                                }
                                returnList.Add((T)instance);
                            }
                        }
                        Xlsxconnection.Close();
                        return returnList;
                    }
                }
            }
            catch (Exception ex)
            {
                logger.LogException(ex, ex.Message);
                return null;
            }
            finally
            {
                CloseStreams();
                if (deleteAfterRead)
                {
                    System.IO.File.Delete(source);
                }
            }
        }
        public object GetConvertedValue(string excelValue, Type type)
        {
            object value = null;
            try
            {
                value = Convert.ChangeType(excelValue, type);
            }
            catch (Exception)
            {
                value = GetDefault(type);
            }
            return value;
        }
        public object GetDefault(Type type)
        {
            if (type.IsValueType)
            {
                return Activator.CreateInstance(type);
            }
            return null;
        }
        private Dictionary<string, string> GetPropertyAttributes<T>()
        {
            Dictionary<string, string> attributesDictionary = new Dictionary<string, string>();
            var properties = typeof(T).GetProperties();
            foreach (PropertyInfo prop in properties)
            {
                object[] attrs = prop.GetCustomAttributes(true);
                foreach (object attr in attrs)
                {
                    ExcelColumnAttribute excelAttribute = attr as ExcelColumnAttribute;
                    if (excelAttribute != null)
                    {
                        string propName = prop.Name;
                        string auth = excelAttribute.Name;
                        attributesDictionary.Add(propName, auth);
                    }
                }
            }
            return attributesDictionary;
        }
        private void CloseStreams()
        {
            if (Xlsxconnection != null && Xlsxconnection.State == ConnectionState.Open)
            {
                Xlsxconnection.Close();
            }
            if (Command != null && Command.Connection != null && Command.Connection.State == ConnectionState.Open)
            {
                Command.Dispose();
            }
            if (Reader != null && !Reader.IsClosed)
            {
                Reader.Close();
            }
        }
    }
}
