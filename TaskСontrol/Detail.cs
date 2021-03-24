using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskСontrol
{
    internal class Detail
    {
        public Detail(string nameDetail, string materialDetail, string thicknessMaterialDetail, string quantityDetail, string quantityDetailNecessary)
        {
            NameDetail = nameDetail ?? throw new ArgumentNullException(nameof(nameDetail));
            MaterialDetail = materialDetail ?? throw new ArgumentNullException(nameof(materialDetail));
            ThicknessMaterialDetail = thicknessMaterialDetail ?? throw new ArgumentNullException(nameof(thicknessMaterialDetail));
            QuantityDetail = quantityDetail ?? throw new ArgumentNullException(nameof(quantityDetail));
            QuantityDetailNecessary = quantityDetailNecessary ?? throw new ArgumentNullException(nameof(quantityDetailNecessary));
        }

        public string NameDetail { get; set; }
        public string MaterialDetail { get; set; }
        public string ThicknessMaterialDetail { get; set; }
        public string QuantityDetail { get; set; }
        public string QuantityDetailNecessary { get; set; }

        public override string ToString()
        {
            return NameDetail + " " + QuantityDetailNecessary + " из " + QuantityDetail + " шт.";
        }
    }
}
