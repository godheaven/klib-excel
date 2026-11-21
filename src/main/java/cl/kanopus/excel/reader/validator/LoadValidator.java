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
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.Date;
import java.util.Locale;
import java.util.Map;
import java.util.regex.Pattern;
import lombok.Getter;
import lombok.Setter;

/**
 * Helpers to validate and convert the values of a row read by {@link cl.kanopus.excel.reader.ExcelReader}.
 *
 * <p>
 * Null/blank handling: a missing key, {@code null} or a blank value is considered "empty". Required validations fail for empty values; non-required parsers return {@code null} for empty values, except the {@code parseString*} methods which return {@code ""} (kept for backward compatibility).
 */
@Setter
@Getter
public class LoadValidator {

    private static final Pattern DOT_GROUPING = Pattern.compile("^\\d{1,3}(\\.\\d{3})+$");

    private boolean titlesToUpperCase = true;

    /**
     * Rounding applied by {@link #parseMoneyToLong(Map, String, boolean)} to amounts with a fractional part. {@code null} (default) rejects amounts with non-zero decimals with {@link ErrorCode#VALUE_NUMBER_FORMAT_EXCEPTION}.
     */
    private RoundingMode moneyRoundingMode;

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

    private String get(Map<String, String> hash, String key) {
        return hash.get(titlesToUpperCase ? key.toUpperCase(Locale.ROOT) : key);
    }

    private static boolean isBlank(String data) {
        return data == null || data.trim().isEmpty();
    }

    private String getRequired(Map<String, String> hash, String key, boolean required) throws LoadValidatorException {
        String data = get(hash, key);
        if (required) {
            validateRequired(data, key);
        }
        return isBlank(data) ? null : data.trim();
    }

    /** Parses a date; returns {@code null} when the value is empty and not required. */
    public Date parseDate(Map<String, String> hash, String key, boolean required, String pattern, int maxLength) throws LoadValidatorException {
        String dateString = parseString(hash, key, required, maxLength);

        Date date = null;
        if (!dateString.isEmpty()) {
            date = Utils.getDate(dateString, pattern);
            if (date == null) {
                throw new LoadValidatorException(ErrorCode.VALUE_DATE_FORMAT_EXCEPTION, key, pattern);
            }
        }
        return date;
    }

    /**
     * Parses a money amount using the Chilean format: {@code "$"} and spaces are ignored, {@code "."} is the thousands separator and {@code ","} the decimal separator (e.g. {@code "$ 1.234.567"} or {@code "1.234,00"}). For values produced by Excel itself, a single {@code "."} not followed by groups of
     * exactly three digits (e.g. {@code "1000.3"}) is read as a decimal separator, and when both separators are present the last one is the decimal separator.
     *
     * <p>
     * Amounts with non-zero decimals are rejected unless {@link #setMoneyRoundingMode(RoundingMode)} is configured. Negative amounts are rejected.
     */
    public Long parseMoneyToLong(Map<String, String> hash, String key, boolean required) throws LoadValidatorException {
        String data = getRequired(hash, key, required);
        if (data == null) {
            return null;
        }
        BigDecimal amount = parseAmount(data.replace("$", "").replace(" ", "").replace("\u00a0", ""), key);
        if (amount.signum() < 0) {
            throw new LoadValidatorException(ErrorCode.VALUE_NUMBER_NEGATIVE_EXCEPTION, key);
        }
        try {
            if (amount.stripTrailingZeros().scale() > 0) {
                if (moneyRoundingMode == null) {
                    throw new LoadValidatorException(ErrorCode.VALUE_NUMBER_FORMAT_EXCEPTION, key);
                }
                amount = amount.setScale(0, moneyRoundingMode);
            }
            return amount.longValueExact();
        } catch (ArithmeticException ae) {
            throw new LoadValidatorException(ErrorCode.VALUE_NUMBER_FORMAT_EXCEPTION, key);
        }
    }

    private static BigDecimal parseAmount(String value, String key) throws LoadValidatorException {
        int lastDot = value.lastIndexOf('.');
        int lastComma = value.lastIndexOf(',');
        String normalized;
        if (lastDot >= 0 && lastComma >= 0) {
            char decimal = lastComma > lastDot ? ',' : '.';
            char grouping = decimal == ',' ? '.' : ',';
            normalized = value.replace(String.valueOf(grouping), "").replace(decimal, '.');
        } else if (lastComma >= 0) {
            normalized = value.indexOf(',') == lastComma ? value.replace(',', '.') : value.replace(",", "");
        } else if (lastDot >= 0) {
            normalized = DOT_GROUPING.matcher(value.startsWith("-") ? value.substring(1) : value).matches() ? value.replace(".", "") : value;
        } else {
            normalized = value;
        }
        try {
            return new BigDecimal(normalized);
        } catch (NumberFormatException ne) {
            throw new LoadValidatorException(ErrorCode.VALUE_NUMBER_FORMAT_EXCEPTION, key);
        }
    }

    public Long parseLong(Map<String, String> hash, String key, boolean required) throws LoadValidatorException {
        String data = getRequired(hash, key, required);
        try {
            return data == null ? null : Long.valueOf(data);
        } catch (NumberFormatException ne) {
            throw new LoadValidatorException(ErrorCode.VALUE_NUMBER_FORMAT_EXCEPTION, key);
        }
    }

    public Integer parseInteger(Map<String, String> hash, String key, boolean required) throws LoadValidatorException {
        String data = getRequired(hash, key, required);
        try {
            return data == null ? null : Integer.valueOf(data);
        } catch (NumberFormatException ne) {
            throw new LoadValidatorException(ErrorCode.VALUE_NUMBER_FORMAT_EXCEPTION, key);
        }
    }

    public String parseStringCutoff(Map<String, String> hash, String key, boolean required, int maxLength) throws LoadValidatorException {
        return parseString(hash, key, required, maxLength, null, true);
    }

    /** Returns the trimmed value, or {@code ""} when it is empty. */
    public String parseString(Map<String, String> hash, String key, boolean required, int maxLength) throws LoadValidatorException {
        return parseString(hash, key, required, maxLength, null, false);
    }

    public String parseString(Map<String, String> hash, String key, boolean required, int maxLength, REGEX regex) throws LoadValidatorException {
        return parseString(hash, key, required, maxLength, regex, false);
    }

    private String parseString(Map<String, String> hash, String key, boolean required, int maxLength, REGEX regex, boolean cutoff) throws LoadValidatorException {
        String data = getRequired(hash, key, required);

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
        return (data == null) ? "" : data;
    }

    /** {@code "true"} and {@code "si"} (ignoring case) are true, other values false; empty values return {@code null}. */
    public Boolean parseBoolean(Map<String, String> hash, String key, boolean required) throws LoadValidatorException {
        String data = getRequired(hash, key, required);
        return data == null ? null : ("true".equalsIgnoreCase(data) || "si".equalsIgnoreCase(data));
    }

    public String parseRut(Map<String, String> hash, String key, boolean required) throws LoadValidatorException {
        String data = parseString(hash, key, required, 100);
        if (!data.isEmpty() && !Utils.isRut(data)) {
            throw new LoadValidatorException(ErrorCode.VALUE_RUT_FORMAT_EXCEPTION, key);
        }
        return !data.isEmpty() ? data.toUpperCase(Locale.ROOT) : null;
    }

    public String parseGTIN(Map<String, String> hash, String key, boolean required) throws LoadValidatorException {
        String data = parseString(hash, key, required, 14);
        if (!data.isEmpty() && !Utils.isGTIN(data)) {
            throw new LoadValidatorException(ErrorCode.VALUE_GTIN_FORMAT_EXCEPTION, key);
        }
        return !data.isEmpty() ? data.toUpperCase(Locale.ROOT) : null;
    }

    /** Fails when the value is {@code null}, empty or blank. Since 4.20.0 {@code "0"} is a valid value. */
    public void validateRequired(String data, String label) throws LoadValidatorException {
        if (isBlank(data)) {
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
