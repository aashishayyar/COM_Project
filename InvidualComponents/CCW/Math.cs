using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ManagementServerForInterop
{
	public interface IMath
	{
		int MultiplicationValue { get; set; }
		int MultiplicationOfTwoIntegers(int iNum1, int iNum2);
		int DivisionOfTwoIntegers(int iNum1, int iNum2);
	}

	[ClassInterface(ClassInterfaceType.AutoDispatch)]
	public class Math : IMath
	{
		public int MultiplicationValue { get; set; }
		public int DivisionValue { get; set; }

		public Math()
		{

		}

		public int MultiplicationOfTwoIntegers(int iNum1, int iNum2)
		{
			MultiplicationValue = iNum1 * iNum2;
			MessageBox.Show("MultiplicationofTwoIntegers : " + MultiplicationValue);
			return MultiplicationValue;
		}

		public int DivisionOfTwoIntegers(int iNum1, int iNum2)
		{
			DivisionValue = iNum1 / iNum2;
			MessageBox.Show("Division of two Integers : " + DivisionValue);
			return DivisionValue;
		}
	} 
}
