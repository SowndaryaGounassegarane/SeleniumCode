package com;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import java.util.Properties;

public class ConfigClass {
  
public static String getProperty(String key)throws FileNotFoundException,IOException{
		{
			Properties p = new Properties();
			FileInputStream f = new FileInputStream(System.getProperty("user.dir")+"\\config\\config.properties");
			p.load(f);
			return p.getProperty(key);
			
		}
		}
    }