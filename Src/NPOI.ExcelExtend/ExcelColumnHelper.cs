using NPOI.ExcelExtend.Attributes;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text;

namespace NPOI.ExcelExtend
{
    public class ExcelColumnHelper
    {
        private List<PropertyInfo> _columns;

        public List<PropertyInfo> ExcelColumns
        {
            get
            {
                if (_columns == null)
                {
                    _columns = GetExcelColumns(ModelType);
                }
                return _columns;
            }
            set { _columns = value; }
        }

        private Type _modelType;

        public Type ModelType
        {
            get { return _modelType; }
            set { _modelType = value; }
        }

        public ExcelColumnHelper(Type type)
        {
            _modelType = type;
        }
        public List<PropertyInfo> GetExcelColumns(Type type)
        {
            var model = new List<PropertyInfo>();
            var propertyInfos = type.GetProperties();

            foreach (var propertyInfo in propertyInfos)
            {
                object[] attrs = propertyInfo.GetCustomAttributes(true);
                foreach (object attr in attrs)
                {
                    ///if this column need to export
                    ExcelColumnAttribute authAttr = attr as ExcelColumnAttribute;
                    if (authAttr != null)
                    {
                        model.Add(propertyInfo);
                    }
                }
            }

            /// order by excel column order prop
            model = model.OrderBy(it => 
            (it.GetCustomAttributes(true).Where(t => t.GetType() == typeof(ExcelColumnAttribute)).Single() as ExcelColumnAttribute).Order).ToList();
            return model;
        }


        /// <summary>
        /// get headers
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public List<string> GetPropertyDisplayNames(List<string> titleList = null, ResourceManager rm = null)
        {
            titleList = (titleList == null ? new List<string>() : titleList);

            foreach (var propertyInfo in ExcelColumns)
            {
                object[] attrs = propertyInfo.GetCustomAttributes(true);
                foreach (object attr in attrs)
                {
                    ///if this column need to export
                    ExcelColumnAttribute authAttr = attr as ExcelColumnAttribute;
                    if (authAttr != null)
                    {
                        var titleName = propertyInfo.GetDisplayName();
                        var displayAttrInfo = propertyInfo.GetCustomAttributes(true).Where(it => it.GetType() == typeof(DisplayAttribute)).SingleOrDefault();
                        if (displayAttrInfo != null)
                        {
                            var displayAttr = displayAttrInfo as DisplayAttribute;
                            var resourceType = displayAttr.ResourceType;
                            if (resourceType != null && rm != null)
                            {
                                //var rm = new ResourceManager(resourceType);
                                titleName = rm.GetString(titleName);
                            }
                        }
                        titleList.Add(titleName);

                    }

                    ///if this object of collection need to export
                    var childDataAttr = attr as ChildDataAttribute;
                    if (childDataAttr != null)
                    {
                        var childHelper = new ExcelColumnHelper(childDataAttr.type);
                        childHelper.GetPropertyDisplayNames(titleList);
                    }
                }

            }

            return titleList;
        }

        /// <summary>
        /// get a martix of object 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <returns></returns>
        public List<List<object>> GetPropertyValues(object data, List<List<object>> allValues = null)
        {
            allValues = allValues == null ? new List<List<object>>() : allValues;
            var propertyValues = new List<object>();
            var propertyInfos = data.GetType().GetProperties();

            foreach (var propertyInfo in ExcelColumns)
            {
                object[] attrs = propertyInfo.GetCustomAttributes(true);
                foreach (object attr in attrs)
                {
                    ExcelColumnAttribute authAttr = attr as ExcelColumnAttribute;
                    if (authAttr != null)
                    {
                        var val = propertyInfo.GetValue(data, null);
                        //var valString = val == null ? "" : val.ToString();
                        propertyValues.Add(val);
                    }
                    ///if this object of collection need to export
                    var childDataAttr = attr as ChildDataAttribute;
                    if (childDataAttr != null)
                    {
                        var vals = propertyInfo.GetValue(data, null);
                        var val = vals as IEnumerable<Object>;
                        foreach (var item in val)
                        {
                            GetPropertyValues(item, allValues);
                        }
                        //GetPropertyDisplayNames()
                    }
                }
            }
            allValues.Insert(0, propertyValues);

            return allValues;
        }
    }
}
