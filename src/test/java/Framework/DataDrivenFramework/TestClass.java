package Framework.DataDrivenFramework;

import java.io.IOException;
import java.util.ArrayList;

public class TestClass {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		ConnectExcel connectExcel = new ConnectExcel();
		ArrayList<String> data = connectExcel.getData("Add Profile");
		System.out.println(data.get(0));
		System.out.println(data.get(1));
		System.out.println(data.get(2));
		System.out.println(data.get(3));
	}

}
