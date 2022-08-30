using System;
using ExcelDna.Integration;
using System.Collections.Generic;
using System.Linq;

namespace PTO
{
    public static partial class PTOfunctionCalls
    {
        [ExcelFunction(Description = "Перевод давления из Паскалей в bar")]
        public static double BAR([ExcelArgument(Name = "P", Description = "Давление, Па")] double Pa)
        {
            return Pa / 100000;
        }
        [ExcelFunction(Description = "Перевод давления из bar в Паскали")]
        public static double ПА([ExcelArgument(Name = "P", Description = "Давление, bar")] double bar)
        {
            return bar * 100000;
        }

        [ExcelFunction(Description = "Перевод энергии из Джоулей в Калории")]
        public static double КАЛ([ExcelArgument(Name = "E", Description = "Энергия, Дж")] double J)
        {
            return J / 4.1868;
        }

        [ExcelFunction(Description = "Перевод энергии из Калорий в Джоули")]
        public static double ДЖ([ExcelArgument(Name = "E", Description = "Энергия, кал")] double K)
        {
            return K * 4.1868;
        }

        [ExcelFunction(Description = "Возвращает температуру насыщения ts, °С")]
        public static double ТНАСЫЩЕНИЯ([ExcelArgument(Name = "Pизб", Description = "Избыточное давление, bar")] double Pi,
                                        [ExcelArgument(Name = "Pатм", Description = "Атмосферное давление, bar\nЕсли не указано применяется 1,01325 bar")] double Patm)
        {
            double Press = CountAbsPress(Pi, Patm);

            return TnasCompute(Press);
        }
        private static double TnasCompute(double Press)
        {
            if (Press >= 221)
                return 374.12;

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
            return Triac(Press, out1.Key, out2.Key, out1.Value, out2.Value);
            //return out1.Value + (out2.Value - out1.Value) / (out2.Key - out1.Key) * (Press - out1.Key);
        }

        [ExcelFunction(Description = "Удельная энтальпия кипящей воды 'h, кДж/кг")]
        public static double ЭНТАЛЬПИЯВ([ExcelArgument(Name = "Pизб", Description = "Избыточное давление, bar")] double Pi,
                                [ExcelArgument(Name = "Pатм", Description = "Атмосферное давление, bar\nЕсли не указано применяется 1,01325 bar")] double Patm)
        {
            double Press = CountAbsPress(Pi, Patm);

            return EntalpyVCompute(Press);
        }
        private static double EntalpyVCompute(double Press)
        {
            KeyValuePair<double, double> out1 = new KeyValuePair<double, double>();
            KeyValuePair<double, double> out2 = new KeyValuePair<double, double>();

            try
            {
                out2 = EntalpyNasV.First(x => x.Key >= Press);
                out1 = EntalpyNasV.Last(x => x.Key <= Press);
            }
            catch (Exception)
            {
                throw new ArgumentOutOfRangeException("Абсолютное давление", Press, "Абсолютное давление больше критического 221,15 бар");
            }

            if (out1.Equals(out2))
                return out1.Value;
            return Triac(Press, out1.Key, out2.Key, out1.Value, out2.Value);
            //return out1.Value + (out2.Value - out1.Value) / (out2.Key - out1.Key) * (Press - out1.Key);
        }

        [ExcelFunction(Description = "Удельная энтальпия сухого насыщенного пара ''h, кДж/кг")]
        public static double ЭНТАЛЬПИЯП([ExcelArgument(Name = "Pизб", Description = "Избыточное давление, bar")] double Pi,
                        [ExcelArgument(Name = "Pатм", Description = "Атмосферное давление, bar\nЕсли не указано применяется 1,01325 bar")] double Patm)
        {
            double Press = CountAbsPress(Pi, Patm);

            return EntalpyPCompute(Press);
        }
        private static double EntalpyPCompute(double Press)
        {
            if (Press < 0)
                throw new ArgumentOutOfRangeException("Абсолютное давление", Press, "Абсолютное давление меньше 0");

            KeyValuePair<double, double> out1 = new KeyValuePair<double, double>();
            KeyValuePair<double, double> out2 = new KeyValuePair<double, double>();

            try
            {
                out2 = EntalpyNasP.First(x => x.Key >= Press);
                out1 = EntalpyNasP.Last(x => x.Key <= Press);
            }
            catch (Exception)
            {
                throw new ArgumentOutOfRangeException("Абсолютное давление", Press, "Абсолютное давление больше критического 221,15 бар");
            }

            if (out1.Equals(out2))
                return out1.Value;
            return Triac(Press, out1.Key, out2.Key, out1.Value, out2.Value);
            //return out1.Value + (out2.Value - out1.Value) / (out2.Key - out1.Key) * (Press - out1.Key);
        }
        private struct Point
        {
            public double T;
            public double P;
            public double E;
            public bool par;
            public int index;
        };
        [ExcelFunction(Description = "Удельная энтальпия h, кДж/кг")]
        public static double ЭНТАЛЬПИЯ([ExcelArgument(Name = "T", Description = "Температура, °C")] double T,
                                        [ExcelArgument(Name = "Pизб", Description = "Избыточное давление, bar")] double Pi,
                                        [ExcelArgument(Name = "Pатм", Description = "Атмосферное давление, bar\nЕсли не указано применяется 1,01325 bar")] double Patm)
        {
            if (T < 0)
                throw new ArgumentOutOfRangeException("Температура", T, "Температура меньше 0 °С");
            if (T > 800)
                throw new ArgumentOutOfRangeException("Температура", T, "Температура больше 800 °С");

            double Press = CountAbsPress(Pi, Patm);

            if (Press < 0)
                throw new ArgumentOutOfRangeException("Абсолютное давление", Press, "Абсолютное давление меньше 0");

            //Находим температуру насыщения для расчетного давления
            double Tn = TnasCompute(Press);

            //Определяем состояние вещества для расчетной точки (вода/пар)
            bool par = T >= Tn;
            Point LT, LB, RT, RB;
            double e1, e2;

            // вычитываем точки
            LT = FindPoint(T - 10.0, Press, -1);
            RT = FindPoint(T - 10.0, Press);
            if (T != 0.0)
            {
                LB = FindPoint(T, Press, -1);
                RB = FindPoint(T, Press);
            }
            else
            {
                LB = FindPoint(T + 10.0, Press, -1);
                RB = FindPoint(T + 10.0, Press);
            }

            //варианты таблицы
            // все 4 точки в том же агрегатном состоянии : все супер - считаем
            if ((LT.par == par && LB.par == par && RT.par == par && RB.par == par) || Press >= 221.15)
            {
                e1 = LT.E + (LB.E - LT.E) / (LB.T - LT.T) * (T - LT.T);
                e2 = RT.E + (RB.E - RT.E) / (RB.T - RT.T) * (T - RT.T);
                return e1 + (e2 - e1) / (RT.P - LT.P) * (Press - LT.P);
            }

            // для текущего состояния - пар
            // левая нижняя пар, остальные вода : левая нижняя из таблицы, левая верхняя точка насыщения, правая верхняя точка насыщения, правая нижняя смещаем вниз
            // левые пар правые вода : левые из таблицы, правый верх точка насыщения, правая нижняя смещаем вниз
            // правая верхняя вода, остальные пар: правая верхняя берем точку насыщения
            // верхние вода, нижние пар : верхние точки насыщения, нижние пар
            // все 4 вода - невозможно
            if (par)
            {
                if (LT.par == false)
                {
                    if (LT.P < 221.15)
                    {
                        LT.T = TnasCompute(LT.P);
                        LT.E = EntalpyPCompute(LT.P);
                    }
                    else
                    {
                        LT.P = 221;
                        LT.T = 374.06;
                        LT.E = 2147.6;
                    }
                }
                if (RT.par == false)
                {
                    if (RT.P < 221.15)
                    {
                        RT.T = TnasCompute(RT.P);
                        RT.E = EntalpyPCompute(RT.P);
                    }
                    else
                    {
                        RT.P = 221;
                        RT.T = 374.06;
                        RT.E = 2147.6;
                    }
                }
                double offset = 10.0;
                while (RB.par == false)
                {
                    RB = FindPoint(T + offset, Press);
                    offset += 10.0;
                }
            }
            // для текущего состояния - вода
            // левая нижняя пар, остальные вода : берем левую нижнюю точку насыщения
            // левые пар правые вода : берем левую нижнюю точку насыщения, левую верхнюю смещаем выше
            // правая верхняя вода, остальные пар: берем левую нижнюю точку насыщения, левую верхнюю смещаем выше, правую нижнюю берем точку насыщения
            // верхние вода, нижние пар : верхние из таблицы нижние точки насыщения
            // все 4 пар - невозможно
            else
            {
                if (LB.par == true)
                {
                    if (LB.P < 221.15)
                    {
                        LB.T = TnasCompute(LB.P);
                        LB.E = EntalpyVCompute(LB.P);
                    }
                    else
                    {
                        LB.P = 221;
                        LB.T = 374.06;
                        LB.E = 2147.6;
                    }
                }
                if (RB.par == true)
                {
                    if (RB.P < 221.15)
                    {
                        RB.T = TnasCompute(RB.P);
                        RB.E = EntalpyVCompute(RB.P);
                    }
                    else
                    {
                        RB.P = 221;
                        RB.T = 374.06;
                        RB.E = 2147.6;
                    }
                }
                double offset = 20.0;
                while (LT.par == true)
                {
                    LT = FindPoint(T - offset, Press, -1);
                    offset -= 10.0;
                }
            }
            e1 = LT.E + (LB.E - LT.E) / (LB.T - LT.T) * (T - LT.T);
            e2 = RT.E + (RB.E - RT.E) / (RB.T - RT.T) * (T - RT.T);
            return e1 + (e2 - e1) / (RB.P - LB.P) * (Press - LB.P);
        }

        [ExcelFunction(Description = "Удельная энтальпия h, ккал/кг")]
        public static double ЭНТАЛЬПИЯПТО([ExcelArgument(Name = "T", Description = "Температура, °C")] double T,
                                        [ExcelArgument(Name = "Pизб", Description = "Избыточное давление, МПа")] double Pi,
                                        [ExcelArgument(Name = "Pатм", Description = "Атмосферное давление, МПа\nЕсли не указано применяется 0,101325 МПа")] double Patm)
        {
            return КАЛ(ЭНТАЛЬПИЯ(T, BAR(Pi * 1000000), BAR(Patm * 1000000)));
        }

        private static Point FindPoint(double T, double P, int Indexoffset = 0)
        {
            Point o = new Point();
            KeyValuePair<double, List<double>> point;

            point = Entalpy.First(x => x.Key >= T);
            o.T = point.Key;
            o.index = EntalpyPressureList.FindIndex(x => x >= P) + Indexoffset;
            o.E = point.Value[o.index];
            o.P = EntalpyPressureList[o.index];
            o.par = o.T >= ТНАСЫЩЕНИЯ(o.P - 1.01325, 0);

            return o;
        }
        private static double Triac(double x, double x1, double x2, double y1, double y2)
        {
            return y1 + (y2 - y1) / (x2 - x1) * (x - x1);
        }
        private static double CountAbsPress(double Pi, double Patm)
        {
            double Press;
            if (Patm == 0)
                Press = Pi + 1.01325;
            else
                Press = Pi + Patm;
            if (Press < 0)
                throw new ArgumentOutOfRangeException("Абсолютное давление", Press, "Абсолютное давление меньше 0");
            return Press;
        }
    }
}
