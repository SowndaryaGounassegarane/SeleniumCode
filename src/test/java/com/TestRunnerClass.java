package com;

import java.util.Collections;

import org.junit.runner.RunWith;
import org.junit.runners.Suite;
import org.testng.TestNG;

import com.example.AOIReportTest;
import com.example.AOIScanTest;

@RunWith(Suite.class)
@Suite.SuiteClasses({
LoginTest.class,
ProductTest.class,
WorkOrderTest.class,
AOIScanTest.class,
AOIReportTest.class,
WaveInputScanTest.class,
WaveOutputScanTest.class,	
WaveReportTest.class,
ReworkTest.class,
ReworkReportTest.class,
MasterReportTest.class,
})
public class TestRunnerClass {
 public static void main(String[] args) {
        TestNG testng = new TestNG();
     testng.setTestSuites(Collections.singletonList("C:\\Users\\sowndarya.g\\VSCode\\SampleTestng\\demo\\testng.xml"));
        testng.run();
    }
}



