using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NUnitSample
{
    public class Property
    {
        //*********************Global locators*******************************************************************************
        public static String BASEURL = "BASEURL";
        // public static String PROPERTY_FILENAME = "config/gui_automation.properties";
        public static String XLS_DATA = "XLS_DATA";
        public static String USERS_LIST = "USERS_LIST";
        public static String ProductMenu_ID = "ProductMenu_ID";
        public static String Product_Management = "Product_Management";
        public static String Specifications = "Specifications";
        public static String Createnewbutton = "Createnewbutton";
        public static String customerSearch = "customerSearch";
        public static String Customer = "Customer";
        public static String SpecId = "SpecId";
        public static String ProductStyle = "ProductStyle";
        public static String GLcode = "GLcode";
        public static String GLCodeSelection = "GLCodeSelection";
        public static String ProductStyleID = "ProductStyleID";
        public static String ProductLength = "ProductLength";
        public static String ProductWidth = "ProductWidth";
        public static String ProductDepth = "ProductDepth";
        public static String MaterialGrade = "MaterialGrade";
        //*[@class='menuIcon menuIcon-Product_Management']
    }
    public class PropertyReader
    {
        public static string GetProperty(string key, ApplicationSettingsBase propertyClass)
        {
            SettingsProperty result = null;
            try
            {
                result = propertyClass.Properties[key];
            }
            catch
            {
            }

            if (result != null)
                return result.DefaultValue.ToString();

            return string.Empty;
        }

        internal static string GetProperty(string shipToDropDown_XPATH, int v)
        {
            throw new NotImplementedException();
        }
    }
}
