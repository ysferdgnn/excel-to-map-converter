package com.seras.testSuites;

import com.seras.interfaces.ExcelParserXSSF;
import com.seras.interfaces.ExcelUtil;
import com.seras.tests.ExcelParserHSSFCoreTest;
import com.seras.tests.ExcelParserXSSFCoreTest;
import com.seras.tests.ExcelToMapConverterTest;
import com.seras.tests.ExcelUtilCoreTest;
import org.junit.runner.RunWith;
import org.junit.runners.Suite;

@RunWith(Suite.class)
@Suite.SuiteClasses(
        {
                ExcelParserHSSFCoreTest.class,
                ExcelParserXSSFCoreTest.class,
                ExcelUtilCoreTest.class,
                ExcelToMapConverterTest.class
        })
public class ExcelParserTestSuite {
}
