using NPOI.SS.UserModel;
using SimpleJSON;
using System;
using System.Collections.Generic;

namespace Excel2Config
{
    public class ExcelScalarValueTypes
    {
        private List<ExcelScalarValue> m_protobufScalarValueTypes;
        private Dictionary<string, ExcelScalarValue> m_mapProtobufScalarValueTypes;

		//Protobuf Scalar Value
		public ExcelScalarValueTypes()
        {
            m_protobufScalarValueTypes = new List<ExcelScalarValue>();

            m_protobufScalarValueTypes.Add(new FloatExcelScalarValue());
            m_protobufScalarValueTypes.Add(new Int32ExcelScalarValue());
            m_protobufScalarValueTypes.Add(new StringExcelScalarValue());
            m_protobufScalarValueTypes.Add(new BoolExcelScalarValue());

            m_mapProtobufScalarValueTypes = new Dictionary<string, ExcelScalarValue>();
            foreach (var item in m_protobufScalarValueTypes)
            {
                m_mapProtobufScalarValueTypes.Add(item.Name, item);
            }
        }


        public bool IsBaseType(string type)
        {
            return m_mapProtobufScalarValueTypes.ContainsKey(type);
        }


        public string ToJson(string type, string[] args)
        {
            if (m_mapProtobufScalarValueTypes.TryGetValue(type, out ExcelScalarValue scalarValue))
            {
                return scalarValue.ToJson(args);
            }
            return string.Empty;
        }

        public string ToJson(string type, string arg)
        {
            if (m_mapProtobufScalarValueTypes.TryGetValue(type, out ExcelScalarValue scalarValue))
            {
                return scalarValue.ToJson(arg);
            }
            return string.Empty;
        }

        public string ToJson(string type, ICell cell)
        {
            if (m_mapProtobufScalarValueTypes.TryGetValue(type, out ExcelScalarValue scalarValue))
            {
                return scalarValue.ToJson(cell);
            }
            return string.Empty;
        }


        public string ToTextProto()
        {
            return string.Empty;
        }

    }


    public abstract class ExcelScalarValue
    {
        public string Name { get; protected set; }

        public abstract string ToJson(string[] args);

        public abstract string ToJson(string arg);

        public abstract string ToJson(ICell cell);
    }

   
   
    
    
}


