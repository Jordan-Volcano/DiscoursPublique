﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DiscoursPublique
{
    public class Bosquejos
    {

        public string[] numeros;
        public string[] datesBosquejo;


        public Bosquejos()
        {


        }


        public Bosquejos(string[] num, string[] dates)
        {
            numeros = num;
            datesBosquejo = dates;

        }

      
    }


    public class Orateur
    {
        public string[] Name;
        public string[] datesBosquejo;


        public Orateur()
        {


        }


        public Orateur(string[] nom, string[] dates)
        {
            Name = nom;
            datesBosquejo = dates;

        }

    }


    public class Frère
        {
            public string Name;
            public string[,] bosquejos;


            public Frère()
            {


            }


            public Frère(string nom, string[] dates)
            {
                Name = nom;
                foreach(string date in dates)
            {

                bosquejos = new string[,] { { date, } };
            }
            

            }


        }
    }
