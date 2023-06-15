package ExcelSample;

import java.io.IOException;

public class ExcelMain {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		ExcelCode.readIntegerData(0, 0);
		ExcelCode.readStringData(0, 1);
		System.out.println(ExcelCode.readIntegerData(0, 0));
		System.out.println(ExcelCode.readStringData(0, 1));

	}

}
