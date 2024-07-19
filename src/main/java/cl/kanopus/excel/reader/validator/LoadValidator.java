package cl.kanopus.excel.reader.validator;

import cl.kanopus.common.util.Utils;
import cl.kanopus.excel.reader.validator.LoadValidatorException.ErrorCode;
import java.util.Date;
import java.util.HashMap;
import java.util.regex.Pattern;

public class LoadValidator {

    private boolean titlesToUpperCase = true;

    public enum REGEX {

        STANDARD_NAME("^([a-zA-Z0-9_ -]+)$"),
        RUT("^([0-9]+-[0-9Kk])$");

        private final String expression;

        REGEX(String expression) {
            this.expression = expression;
        }

        public String getExpression() {
            return expression;
        }

    }

    public boolean isTitlesToUpperCase() {
        return titlesToUpperCase;
    }

    public void setTitlesToUpperCase(boolean titlesToUpperCase) {
        this.titlesToUpperCase = titlesToUpperCase;
    }

    public Date parseDate(HashMap<String, String> hash, String key, boolean required, String pattern, int maxlength) throws LoadValidatorException {
        String dateString = parseString(hash, key, required, maxlength);

        Date date = null;
        if (dateString != null) {
            date = Utils.getDate(dateString, pattern);
            if (date == null) {
                throw new LoadValidatorException(ErrorCode.VALUE_DATE_FORMAT_EXCEPTION, key, pattern);
            }
        }
        return date;
    }

    public Long parseMoneyToLong(HashMap<String, String> hash, String key, boolean required) throws LoadValidatorException {
        String data = hash.get(titlesToUpperCase ? key.toUpperCase() : key);
        if (required) {
            validateRequired(data, key);
        }
        try {
            return (data == null || "".equals(data)) ? null : Long.valueOf(data.replaceAll("\\$", "").replaceAll("\\.", "").replaceAll(",", "").trim());
        } catch (NumberFormatException ne) {
            throw new LoadValidatorException(ErrorCode.VALUE_NUMBER_FORMAT_EXCEPTION, key);
        }
    }

    public Long parseLong(HashMap<String, String> hash, String key, boolean required) throws LoadValidatorException {
        String data = hash.get(titlesToUpperCase ? key.toUpperCase() : key);
        if (required) {
            validateRequired(data, key);
        }
        try {
            return (data == null || "".equals(data)) ? null : Long.valueOf(data);
        } catch (NumberFormatException ne) {
            throw new LoadValidatorException(ErrorCode.VALUE_NUMBER_FORMAT_EXCEPTION, key);
        }
    }

    public Integer parseInteger(HashMap<String, String> hash, String key, boolean required) throws LoadValidatorException {
        String data = hash.get(titlesToUpperCase ? key.toUpperCase() : key);
        if (required) {
            validateRequired(data, key);
        }
        try {
            return (data == null || "".equals(data)) ? null : Integer.valueOf(data);
        } catch (NumberFormatException ne) {
            throw new LoadValidatorException(ErrorCode.VALUE_NUMBER_FORMAT_EXCEPTION, key);
        }
    }

    public String parseString(HashMap<String, String> hash, String key, boolean required, int maxLength, REGEX regex) throws LoadValidatorException {
        String data = hash.get(titlesToUpperCase ? key.toUpperCase() : key);
        if (required) {
            validateRequired(data, key);
        }
        validateMaxlength(data, maxLength, key);
        if (regex != null) {
            validaRegex(data, regex, key);
        }
        return (data == null) ? "" : data.trim();
    }

    public boolean parseBoolean(HashMap<String, String> hash, String key, boolean required) throws LoadValidatorException {
        String data = hash.get(titlesToUpperCase ? key.toUpperCase() : key);
        if (required) {
            validateRequired(data, key);
        }

        return (data != null && ("true".equals(data.toLowerCase()) || "si".equals(data.toLowerCase())));
    }

    public String parseString(HashMap<String, String> hash, String key, boolean required, int maxLength) throws LoadValidatorException {
        return parseString(hash, key, required, maxLength, null);
    }

    public String parseRut(HashMap<String, String> hash, String key, boolean required) throws LoadValidatorException {
        String data = parseString(hash, key, required, 10, REGEX.RUT);
        if (data != null && !Utils.isRut(data)) {
            throw new LoadValidatorException(ErrorCode.VALUE_RUT_FORMAT_EXCEPTION, key);
        }
        return data != null ? data.toUpperCase() : null;
    }

    public void validateRequired(String data, String label) throws LoadValidatorException {
        if (data == null || "".equals(data) || "0".equals(data)) {
            throw new LoadValidatorException(ErrorCode.VALUE_REQUIRED, label);
        }
    }

    public void validaRegex(String data, REGEX regex, String label) throws LoadValidatorException {
        if (data != null && !"".equals(data) && !Pattern.matches(regex.getExpression(), data)) {
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
