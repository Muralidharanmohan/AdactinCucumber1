package org.com;

import java.io.IOException;

public class FacebookLogin extends BaseClass {

	public static void main(String[] args) throws IOException {

		getDriver();
		loadurl("https://www.facebook.com/");
		maximize();

	}

}
