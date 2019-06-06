using System;
using System.Collections.Generic;
using System.Text;

namespace Examples
{
    public class People
    {
        public string Name { get; set; }

        public int Age => (DateTime.Today.Year - Birthday.Year + 1);

        public string Address { get; set; }

        public DateTime Birthday { get; set; }

        public string Education { get; set; }

        public bool hasWork { get; set; }

        public string Remark { get; set; }


    }
}
