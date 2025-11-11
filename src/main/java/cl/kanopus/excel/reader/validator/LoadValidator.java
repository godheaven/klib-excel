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

import cl.kanopus.common.util.Utils;
import cl.kanopus.excel.reader.validator.LoadValidatorException.ErrorCode;
import lombok.Getter;
import lombok.Setter;

import java.util.Date;
import java.util.Map;
import java.util.regex.Pattern;

@Setter
@Getter
public class LoadValidator {
    
    private boolean titlesToUpperCase = true;

    @Getter
    public enum REGEX {

        STANDARD_NAME("^([a-zA-Z0-9_ -]+)$"),
        RUT("^([0-9]+-[0-9Kk])$"),
        NUMBER("^([0-9]+)$");

        private final String expression;

        REGEX(String expression) {
            this.expression = expression;
        }

    }

    public Date parseDate(Map<String, String> hash, String key, boolean required, String pattern, int maxLength) throws LoadValidatorException {
        String dateString = parseString(hash, key, required, maxLength);

        Date date = null;
        if (dateString != null) {
            date = Utils.getDate(dateString, pattern);
            if (date == null) {
                throw new LoadValidatorException(ErrorCode.VALUE_DATE_FORMAT_EXCEPTION, key, pattern);
            }
        }
        return date;
    }

    public Long parseMoneyToLong(Map<String, String> hash, String key, boolean required) throws LoadValidatorException {
        String data = hash.get(titlesToUpperCase ? key.toUpperCase() : key);
        if (required) {
            validateRequired(data, key);
        }
        try {
            Long value = null;
            if (!Utils.isNullOrEmpty(data)) {
                value = Long.valueOf(data.replaceAll("\\$", "").replaceAll("\\.", "").replaceAll(",", "").trim());
                if (value < 0) {
                    throw new LoadValidatorException(ErrorCode.VALUE_NUMBER_NEGATIVE_EXCEPTION, key);
                }
            }
            return value;
        } catch (NumberFormatException ne) {
            throw new LoadValidatorException(ErrorCode.VALUE_NUMBER_FORMAT_EXCEPTION, key);
        }
    }

    public Long parseLong(Map<String, String> hash, String key, boolean required) throws LoadValidatorException {
        String data = hash.get(titlesToUpperCase ? key.toUpperCase() : key);
        if (required) {
            validateRequired(data, key);
        }
        try {
            return (data == null || data.isEmpty()) ? null : Long.valueOf(data);
        } catch (NumberFormatException ne) {
            throw new LoadValidatorException(ErrorCode.VALUE_NUMBER_FORMAT_EXCEPTION, key);
        }
    }

    public Integer parseInteger(Map<String, String> hash, String key, boolean required) throws LoadValidatorException {
        String data = hash.get(titlesToUpperCase ? key.toUpperCase() : key);
        if (required) {
            validateRequired(data, key);
        }
        try {
            return (data == null || data.isEmpty()) ? null : Integer.valueOf(data);
        } catch (NumberFormatException ne) {
            throw new LoadValidatorException(ErrorCode.VALUE_NUMBER_FORMAT_EXCEPTION, key);
        }
    }

    public String parseStringCutoff(Map<String, String> hash, String key, boolean required, int maxLength) throws LoadValidatorException {
        return parseString(hash, key, required, maxLength, null, true);
    }

    public String parseString(Map<String, String> hash, String key, boolean required, int maxLength) throws LoadValidatorException {
        return parseString(hash, key, required, maxLength, null, false);
    }

    public String parseString(Map<String, String> hash, String key, boolean required, int maxLength, REGEX regex) throws LoadValidatorException {
        return parseString(hash, key, required, maxLength, regex, false);
    }

    private String parseString(Map<String, String> hash, String key, boolean required, int maxLength, REGEX regex, boolean cutoff) throws LoadValidatorException {
        String data = hash.get(titlesToUpperCase ? key.toUpperCase() : key);
        if (required) {
            validateRequired(data, key);
        }

        if (cutoff) {
            if (data != null && data.length() > maxLength) {
                data = data.substring(0, maxLength);
            }
        } else {
            validateMaxlength(data, maxLength, key);
        }

        if (regex != null) {
            validaRegex(data, regex, key);
        }
        return (data == null) ? "" : data.trim();
    }

    public Boolean parseBoolean(Map<String, String> hash, String key, boolean required) throws LoadValidatorException {
        String data = hash.get(titlesToUpperCase ? key.toUpperCase() : key);
        if (required) {
            validateRequired(data, key);
        }

        return data == null ? null : ("true".equalsIgnoreCase(data) || "si".equalsIgnoreCase(data));
    }

    public String parseRut(Map<String, String> hash, String key, boolean required) throws LoadValidatorException {
        String data = parseString(hash, key, required, 100);
        if (!Utils.isNullOrEmpty(data) && !Utils.isRut(data)) {
            throw new LoadValidatorException(ErrorCode.VALUE_RUT_FORMAT_EXCEPTION, key);
        }
        return !Utils.isNullOrEmpty(data) ? data.toUpperCase() : null;
    }

    public String parseGTIN(Map<String, String> hash, String key, boolean required) throws LoadValidatorException {
        String data = parseString(hash, key, required, 14);
        if (!Utils.isNullOrEmpty(data) && !Utils.isGTIN(data)) {
            throw new LoadValidatorException(ErrorCode.VALUE_GTIN_FORMAT_EXCEPTION, key);
        }
        return !Utils.isNullOrEmpty(data) ? data.toUpperCase() : null;
    }

    public void validateRequired(String data, String label) throws LoadValidatorException {
        if (data == null || data.isEmpty() || "0".equals(data)) {
            throw new LoadValidatorException(ErrorCode.VALUE_REQUIRED, label);
        }
    }

    public void validaRegex(String data, REGEX regex, String label) throws LoadValidatorException {
        if (data != null && !data.isEmpty() && !Pattern.matches(regex.getExpression(), data)) {
            throw new LoadValidatorException(ErrorCode.VALUE_VIOLATES_REGEX, label, regex.getExpression());
        }
    }

    public void validateMaxlength(String data, int maxLength, String label) throws LoadValidatorException {
        if (data != null && data.length() > maxLength) {
            throw new LoadValidatorException(ErrorCode.VALUE_VIOLATES_MAXLENGTH, label, maxLength);
        }
    }

    public void validateRecordNotNull(Object object, String label) throws LoadValidatorException {
        if (object == null) {
            throw new LoadValidatorException(ErrorCode.RECORD_DOES_NOT_EXIST_IN_DATABASE, label);
        }
    }
}
