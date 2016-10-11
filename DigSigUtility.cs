using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DigSigOOXml
{

    class DigSigUtility
    {
        public string _top { get; set; }
        public string _left { get; set; }
        public string _width { get; set; }
        public string _height { get; set; }
        public string _docWidth;
        public string _docHeight;
        public string label { get; set; }
        public string page { get; set; }
        public string name { get; set; }
        private string bottom;

        public DigSigUtility()
        {
            //default to pdf values    
        }

        public DigSigUtility(string width, string height)
        {
            { _docWidth = width;}
            { _docHeight = height; }
        }

        public void setDocWidth(string value)
        {
            _docWidth = calculate(value, 612);
        }

        public void setDocHeight(string value)
        {
            _docHeight = calculate(value, 792);
        }
        public string command()
        {

           
            string s = String.Format("name={0}&label={1}&page={2}&type=formfield&subtype=signature&bottom={3}&left={4}&width={5}&height={6}", name, label, page, calculateBottom(), Convert.ToInt32(Convert.ToDecimal(_left)), Convert.ToInt32(Convert.ToDecimal(_width)), Convert.ToInt32(Convert.ToDecimal(_height))); 
            return s;
        }

        private string calculateBottom()
        {

            decimal btmValue = 0;

            if (_top != null && _height != null)
            {
                decimal h = Convert.ToDecimal(_height);
                decimal t = Convert.ToDecimal(_top);
                decimal i = h + t;
                decimal dh = Convert.ToDecimal(_docHeight);
                btmValue = dh - i;


            }

            return Convert.ToString(Convert.ToInt32(btmValue));


        }

        private string calculate(string dim, int std)
        {
            decimal convertedValue = std;

            if (dim != null)
            {
                decimal k = Convert.ToDecimal(dim);
                if (k > 0)
                {
                    convertedValue = k / 20;
                }
            }

            return Convert.ToString(Convert.ToInt32(convertedValue));
        }

      
    
    }
}
