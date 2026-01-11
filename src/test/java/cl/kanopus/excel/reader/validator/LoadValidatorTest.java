/*-
 * !--
 * For support and inquiries regarding this library, please contact:
 *   soporte@kanopus.cl
 *
 * Project website:
 *   https://www.kanopus.cl
 * %%
 * Copyright (C) 2025 Pablo Díaz Saavedra
 * %%
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * --!
 */
package cl.kanopus.excel.reader.validator;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

class LoadValidatorTest {

    /**
     * Test of isTitlesToUpperCase method, of class LoadValidator.
     */
    @Test
    void testIsTitlesToUpperCase() {
        LoadValidator validator = new LoadValidator();
        // default true
        Assertions.assertTrue(validator.isTitlesToUpperCase());
        validator.setTitlesToUpperCase(false);
        Assertions.assertFalse(validator.isTitlesToUpperCase());
    }

    /**
     * Test of setTitlesToUpperCase method, of class LoadValidator.
     */
    @Test
    void testSetTitlesToUpperCase() {
        LoadValidator validator = new LoadValidator();
        validator.setTitlesToUpperCase(false);
        Assertions.assertFalse(validator.isTitlesToUpperCase());
        validator.setTitlesToUpperCase(true);
        Assertions.assertTrue(validator.isTitlesToUpperCase());
    }

    /**
     * Test of parseDate method, of class LoadValidator.
     */
    @Test
    void testParseDate() throws Exception {
        Map<String, String> hash = new HashMap<>();
        // date in format dd-MM-yyyy
        hash.put("fecha", "11-01-2016");

        LoadValidator validator = new LoadValidator();
        validator.setTitlesToUpperCase(false);
        // pattern expecting dd-MM-yyyy
        Date d = validator.parseDate(hash, "fecha", true, "dd-MM-yyyy", 10);
        SimpleDateFormat df = new SimpleDateFormat("dd-MM-yyyy");
        Assertions.assertEquals("11-01-2016", df.format(d));

        // invalid format should throw
        hash.put("fecha2", "2016/01/11");
        LoadValidatorException ex = Assertions.assertThrows(LoadValidatorException.class, () -> {
            validator.parseDate(hash, "fecha2", true, "dd-MM-yyyy", 10);
        });
        Assertions.assertEquals(LoadValidatorException.ErrorCode.VALUE_DATE_FORMAT_EXCEPTION, ex.getCode());
    }

    /**
     * Test of parseMoneyToLong method, of class LoadValidator.
     */
    @Test
    void testParseMoneyToLong() throws Exception {
        Map<String, String> hash = new HashMap<>();
        hash.put("data1", "-1000");

        hash.put("data2", "1.000");
        hash.put("data3", "1000");
        hash.put("data4", "1000.3");

        LoadValidator validator = new LoadValidator();
        validator.setTitlesToUpperCase(false);
        LoadValidatorException exception1 = Assertions.assertThrows(LoadValidatorException.class, () -> validator.parseMoneyToLong(hash, "data1", true));
        Assertions.assertEquals(LoadValidatorException.ErrorCode.VALUE_NUMBER_NEGATIVE_EXCEPTION, exception1.getCode());

        Assertions.assertEquals(Long.valueOf(1000), validator.parseMoneyToLong(hash, "data2", true));
        Assertions.assertEquals(Long.valueOf(1000), validator.parseMoneyToLong(hash, "data3", true));

    }

    /**
     * Test of parseLong method, of class LoadValidator.
     */
    @Test
    void testParseLong() throws Exception {
        Map<String, String> hash = new HashMap<>();
        hash.put("l1", "12345");
        hash.put("l2", "");
        hash.put("l3", "abc");

        LoadValidator validator = new LoadValidator();
        validator.setTitlesToUpperCase(false);
        Assertions.assertEquals(Long.valueOf(12345L), validator.parseLong(hash, "l1", true));
        Assertions.assertNull(validator.parseLong(hash, "l2", false));
        LoadValidatorException ex = Assertions.assertThrows(LoadValidatorException.class, () -> validator.parseLong(hash, "l3", true));
        Assertions.assertEquals(LoadValidatorException.ErrorCode.VALUE_NUMBER_FORMAT_EXCEPTION, ex.getCode());
    }

    /**
     * Test of parseInteger method, of class LoadValidator.
     */
    @Test
    void testParseInteger() throws Exception {
        Map<String, String> hash = new HashMap<>();
        hash.put("i1", "123");
        hash.put("i2", "");
        hash.put("i3", "1.2");

        LoadValidator validator = new LoadValidator();
        validator.setTitlesToUpperCase(false);
        Assertions.assertEquals(Integer.valueOf(123), validator.parseInteger(hash, "i1", true));
        Assertions.assertNull(validator.parseInteger(hash, "i2", false));
        LoadValidatorException ex = Assertions.assertThrows(LoadValidatorException.class, () -> validator.parseInteger(hash, "i3", true));
        Assertions.assertEquals(LoadValidatorException.ErrorCode.VALUE_NUMBER_FORMAT_EXCEPTION, ex.getCode());
    }

    /**
     * Test of parseString method, of class LoadValidator.
     */
    @Test
    void testParseString_5args() throws Exception {
        Map<String, String> hash = new HashMap<>();
        hash.put("data1", "123456");

        LoadValidator validator = new LoadValidator();
        validator.setTitlesToUpperCase(false);
        Assertions.assertEquals("12345", validator.parseStringCutoff(hash, "data1", true, 5));
    }

    /**
     * Test of parseBoolean method, of class LoadValidator.
     */
    @Test
    void testParseBoolean() throws Exception {
        Map<String, String> hash = new HashMap<>();
        hash.put("b1", "true");
        hash.put("b2", "Si");
        hash.put("b3", "no");

        LoadValidator validator = new LoadValidator();
        validator.setTitlesToUpperCase(false);
        Assertions.assertTrue(validator.parseBoolean(hash, "b1", true));
        Assertions.assertTrue(validator.parseBoolean(hash, "b2", true));
        Assertions.assertFalse(validator.parseBoolean(hash, "b3", true));
    }

    /**
     * Test of parseString method, of class LoadValidator.
     */
    @Test
    void testParseString_4args() throws Exception {
        Map<String, String> hash = new HashMap<>();
        hash.put("s1", " hello ");
        LoadValidator validator = new LoadValidator();
        validator.setTitlesToUpperCase(false);
        Assertions.assertEquals("hello", validator.parseString(hash, "s1", true, 10));

        // null/empty handling
        hash.put("s2", "");
        Assertions.assertEquals("", validator.parseString(hash, "s2", false, 10));

        // max length violation
        hash.put("s3", "123456");
        LoadValidatorException ex = Assertions.assertThrows(LoadValidatorException.class, () -> validator.parseString(hash, "s3", true, 3));
        Assertions.assertEquals(LoadValidatorException.ErrorCode.VALUE_VIOLATES_MAXLENGTH, ex.getCode());
    }

    /**
     * Test of parseRut method, of class LoadValidator.
     */
    @Test
    void testParseRut() throws Exception {
        Map<String, String> hash = new HashMap<>();
        hash.put("rut1", "12345678-5");
        hash.put("rut2", "invalid-rut");

        LoadValidator validator = new LoadValidator();
        validator.setTitlesToUpperCase(false);
        Assertions.assertEquals("12345678-5", validator.parseRut(hash, "rut1", true));
        LoadValidatorException ex = Assertions.assertThrows(LoadValidatorException.class, () -> validator.parseRut(hash, "rut2", true));
        Assertions.assertEquals(LoadValidatorException.ErrorCode.VALUE_RUT_FORMAT_EXCEPTION, ex.getCode());
    }

    /**
     * Test of parseGTIN method, of class LoadValidator.
     */
    @Test
    void testParseGTIN() throws Exception {
        Map<String, String> hash = new HashMap<>();
        hash.put("g1", "0000337787551");
        hash.put("g2", "abcd");

        LoadValidator validator = new LoadValidator();
        validator.setTitlesToUpperCase(false);
        // depending on Utils.isGTIN implementation, valid numeric string of length <=14 should pass
        Assertions.assertEquals("0000337787551", validator.parseGTIN(hash, "g1", true));
        LoadValidatorException ex = Assertions.assertThrows(LoadValidatorException.class, () -> validator.parseGTIN(hash, "g2", true));
        Assertions.assertEquals(LoadValidatorException.ErrorCode.VALUE_GTIN_FORMAT_EXCEPTION, ex.getCode());
    }

    /**
     * Test of validateRequired method, of class LoadValidator.
     */
    @Test
    void testValidateRequired() {
        LoadValidator validator = new LoadValidator();
        LoadValidatorException ex = Assertions.assertThrows(LoadValidatorException.class, () -> validator.validateRequired(null, "label"));
        Assertions.assertEquals(LoadValidatorException.ErrorCode.VALUE_REQUIRED, ex.getCode());

        ex = Assertions.assertThrows(LoadValidatorException.class, () -> validator.validateRequired("", "label"));
        Assertions.assertEquals(LoadValidatorException.ErrorCode.VALUE_REQUIRED, ex.getCode());

        // '0' is considered required fail per impl
        ex = Assertions.assertThrows(LoadValidatorException.class, () -> validator.validateRequired("0", "label"));
        Assertions.assertEquals(LoadValidatorException.ErrorCode.VALUE_REQUIRED, ex.getCode());
    }

    /**
     * Test of validaRegex method, of class LoadValidator.
     */
    @Test
    void testValidaRegex() throws Exception {
        LoadValidator validator = new LoadValidator();
        // valid
        validator.validaRegex("ABCD-123", LoadValidator.REGEX.STANDARD_NAME, "lbl");
        // invalid
        LoadValidatorException ex = Assertions.assertThrows(LoadValidatorException.class, () -> validator.validaRegex("$%$%", LoadValidator.REGEX.STANDARD_NAME, "lbl"));
        Assertions.assertEquals(LoadValidatorException.ErrorCode.VALUE_VIOLATES_REGEX, ex.getCode());
    }

    /**
     * Test of validateMax length method, of class LoadValidator.
     */
    @Test
    void testValidateMaxLength() throws Exception {
        LoadValidator validator = new LoadValidator();
        // no exception
        validator.validateMaxlength("abc", 5, "lbl");
        // exception when longer
        LoadValidatorException ex = Assertions.assertThrows(LoadValidatorException.class, () -> validator.validateMaxlength("abcdef", 5, "lbl"));
        Assertions.assertEquals(LoadValidatorException.ErrorCode.VALUE_VIOLATES_MAXLENGTH, ex.getCode());
    }

    /**
     * Test of validateRecordNotNull method, of class LoadValidator.
     */
    @Test
    void testValidateRecordNotNull() throws Exception {
        LoadValidator validator = new LoadValidator();
        LoadValidatorException ex = Assertions.assertThrows(LoadValidatorException.class, () -> validator.validateRecordNotNull(null, "lbl"));
        Assertions.assertEquals(LoadValidatorException.ErrorCode.RECORD_DOES_NOT_EXIST_IN_DATABASE, ex.getCode());

        // no exception when object present
        validator.validateRecordNotNull(new Object(), "lbl");
    }

}
