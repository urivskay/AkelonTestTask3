using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Security.Permissions;
using System.Text;
using System.Threading.Tasks;

namespace AkelonTestExcel.Repository
{
    internal class Product
    {
        public int Article { get; set; }
        public string Name { get; set; }
        public Units Units { get; set; }
        public float Price { get; set; }

        public static IEnumerable<string> GetDescriptionsUnits()
        {
            var descs = new List<string>();
            var names = Enum.GetNames(typeof(Units));
            foreach (var name in names)
            {
                var field = typeof(Units).GetField(name);
                var fds = field?.GetCustomAttributes(typeof(DescriptionAttribute), true);

                if (fds != null)
                    foreach (DescriptionAttribute fd in fds)
                    {
                        descs.Add(fd.Description);
                    }
            }
            return descs;
        }
    }

    public enum Units
    {
        NoUnits = 0,
        [Description("Килограмм")]
        Kg = 1,
        [Description("Литр")]
        Liter = 2,
        [Description("Штука")]
        Count = 3
    }
}
