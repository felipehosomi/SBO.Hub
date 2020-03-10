using SBO.Hub.DAO;
using SBO.Hub.SBOModels;
using System;
using System.Collections.Generic;

namespace SBO.Hub.SBOObjects
{
    public class City : SystemObjectDAO
    {
        public City()
            : base("OCNT")
        { }

        public List<CityModel> GetCitiesList(string country, string uf)
        {
            string where = "Country = '{0}' AND State = '{1}'";
            where = String.Format(where, country, uf);

            List<CityModel> list = this.RetrieveModelList<CityModel>(where);
            return list;
        }
    }
}
