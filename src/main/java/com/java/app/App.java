package com.java.app;

import java.util.Arrays;
import java.util.concurrent.ThreadLocalRandom;

import com.java.app.excelprocessor.ReadExcelFileToList;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
        System.out.println( "Hello World! Java Starter Project Here!");
        ReadExcelFileToList.main(new String[0]);
     }
}
