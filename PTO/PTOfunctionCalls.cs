﻿using System;
using ExcelDna.Integration;
using System.Collections.Generic;
using System.Linq;

namespace PTO
{
    public static partial class PTOfunctionCalls
    {
        [ExcelFunction(Description = "Перевод давления из Паскалей в bar")]
        public static double BAR([ExcelArgument(Name = "P", Description = "Давление в Паскалях")] double Pa)
        {
            return Pa / 100000;
        }

        [ExcelFunction(Description = "Возвращает температуру насыщения, °С")]
        public static double ТНАСЫЩЕНИЯ([ExcelArgument(Name = "Pизб", Description = "Избыточное давление, bar")] double Pa,
                                        [ExcelArgument(Name = "Pатм", Description = "Атмосферное давление, bar\nЕсли не указано применяется 1,01325 bar")] double Patm)
        {
            double Press = 0;

            if (Patm == 0)
                Press = Pa + 1.01325;
            else
                Press = Pa + Patm;
            if (Press < 0)
                throw new ArgumentOutOfRangeException("Абсолютное давление", Press, "Абсолютное давление меньше 0");

            KeyValuePair<double, double> out1 = new KeyValuePair<double, double>();
            KeyValuePair<double, double> out2 = new KeyValuePair<double, double>();

            try
            {
                out2 = Tnas.First(x => x.Key >= Press);
                out1 = Tnas.Last(x => x.Key <= Press);
            }
            catch (Exception)
            {
                throw new ArgumentOutOfRangeException("Абсолютное давление", Press, "Абсолютное давление больше критического 221,15 бар");
            }

            if (out1.Equals(out2))
                return out1.Value;
            return out1.Value + (out2.Value - out1.Value) / (out2.Key - out1.Key) * (Press - out1.Key);
        }

        [ExcelFunction(Description = "Возвращает температуру насыщения, °С")]
        public static double ЭНТАЛЬПИЯ([ExcelArgument(Name = "T", Description = "Температура, °C")] double T,
                                        [ExcelArgument(Name = "Pизб", Description = "Избыточное давление, bar")] double Pa,
                                        [ExcelArgument(Name = "Pатм", Description = "Атмосферное давление, bar\nЕсли не указано применяется 1,01325 bar")] double Patm)
        {
            double Press = 0;

            if (T < 0)
                throw new ArgumentOutOfRangeException("Температура", T, "Температура меньше 0");
            if (Patm == 0)
                Press = Pa + 1.01325;
            else
                Press = Pa + Patm;
            if (Press < 0)
                throw new ArgumentOutOfRangeException("Абсолютное давление", Press, "Абсолютное давление меньше 0");

            double Tn = ТНАСЫЩЕНИЯ(Pa, Patm);



            return 0;
        }
    }
}