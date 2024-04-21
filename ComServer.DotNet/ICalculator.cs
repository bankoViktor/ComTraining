using System;
using System.Runtime.InteropServices;

namespace ComServer.DotNet
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("7FA6BD16-F8C8-4F86-A240-E5BDB4D254EF")]
    public interface ICalculator
    {
        double Sum(double a, double b);
        double Sub(double a, double b);
        double Mul(double a, double b);
        double Div(double a, double b);
    }
}
