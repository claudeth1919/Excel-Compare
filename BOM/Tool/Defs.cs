using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BOM.Tool
{
    public enum AlarmType{
        SUCCESS,
        ERROR,
        WARNING,
        LOADING
    }

    public enum Source
    {
        FILE_1,
        FILE_2
    }
    public static class Defs
    {
        public static String DB_HOSTNAME= "mx020038";
        public static String DB_DATABASE= "BOM";
        public static String DB_USER= "AIS_ALDS";
        public static String DB_PASSWORD= "23erghiop0";
        public static String COL_MATERIAL_CODE = "CLAVE";
        public static String COL_AMOUNT = "CANTIDAD";
        public static List<string> COL_NUMBER_ARTICULE_POSSIBLE_LIST = new List<string>() { "CLAVE", "NUMERO", "NÚMERO" };
        public static List<string> COL_AMOUNT_ARTICULE_POSSIBLE_LIST = new List<string>() { "CANTIDAD", "AMOUNT" };

        //Excel
        public static String EXCEL_FILE_POSTFIX = "_REVISADO.xlsx";

        //Begin Avila
        public static String COL_MATERIAL_ORDERN_AVILA = "PARTE";
        public static String COL_DESCRIPCION_AVILA = "DESCRIPCION";
        public static String COL_MARCA_AVILA = "PROVEEDOR";
        public static String COL_PZAS_AVILA = "CANTIDAD";
        //End Avila


        //Begin Memo
        public static String COL_MATERIAL_ORDERN_MEMO = "NUMERO";
        public static String COL_DESCRIPCION_MEMO = "DESCRIPCION";
        public static String COL_MARCA_MEMO = "PROVEEDOR";
        public static String COL_PZAS_MEMO = "CANTIDAD";
        //End Memo

        public static String COL_WATERHOUSE_IDS = "ARTICULO";
        public static String COL_LOCATION = "UBICACIN";
        public static String COL_PROVIDER = "FABRICANTE";
        public static String COL_TOTAL = "TOTAL";
        public static String COL_KLT = "KLT";
        public static String COL_UNIT = "MEDICI";
        public static String COL_STATUS = "STATUS";

        public static String SHEET_NAME = "SHEET";
        public static String EXCEL_ROW = "ROW";
        public static String EXCEL_CONSUMIBLE_SHEET = "CONSUMIBLE";


        


        public static String EXCEL_TAKEN_STATUS = "RESERVADO";

        public static int NUM_COL_MATERIAL_CODE = 1;
        public static int NUM_COL_AMOUNT = 2;


        public static int ORIGIN_PROJECT_LEFTOVERS = 4;
        public static int ORIGIN_PONTON_WAREHOUSE = 1;

        public static string COLOR_GREEN = "ff92d050";

        public static String EXCLUDED_SHEET_COSTOS = "COSTOS";
        public static String EXCLUDED_SHEET_LEYENDA = "LEYENDA";
        public static String EXCLUDED_SHEET_COLORES = "COLORES";


        public static String STATE_PENDING = "PENDIENTE";
        public static String STATE_ORIGINAL = "ORIGINAL";
        public static String STATE_RESERVED = "RESERVADO";


        public static String NONE_COL = "NONE(N/A)";
        public static int NONE_INDEX_COL = 100000000;

        public static String ERROR_CLOSEN_FILE = "The object invoked has disconnected from its clients";

    }
}
