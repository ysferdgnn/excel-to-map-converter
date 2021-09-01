package com.seras.testSuites;

import com.seras.tests.ExcelParserHSSFCoreTest;
import com.seras.tests.ExcelParserXSSFCoreTest;
import com.seras.tests.ExcelToMapConverterImplementationTest;
import com.seras.tests.ExcelUtilCoreTest;
import org.junit.runner.RunWith;
import org.junit.runners.Suite;

@RunWith(Suite.class)
@Suite.SuiteClasses(
        {
                ExcelParserHSSFCoreTest.class,
                ExcelParserXSSFCoreTest.class,
                ExcelUtilCoreTest.class,
                ExcelToMapConverterImplementationTest.class
        })
public class ExcelParserTestSuite {
}
