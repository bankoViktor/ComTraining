using System;
using System.Runtime.InteropServices;

namespace ComServer.DotNet
{
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("921F2D2C-9752-4BC7-AB2E-E5FB39D165B5")]
    public class CoCalculator : ICalculator
    {
        public double Sum(double a, double b)
        {
            return a + b;
        }
        public double Sub(double a, double b)
        {
            return a - b;
        }

        public double Div(double a, double b)
        {
            if (b == 0)
                throw new Exception("Can't decide on 0.");
            return a / b;
        }

        public double Mul(double a, double b)
        {
            return a * b;
        }
    }
}
