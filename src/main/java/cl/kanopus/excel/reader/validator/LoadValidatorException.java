package cl.kanopus.excel.reader.validator;

import cl.kanopus.common.util.Utils;

public class LoadValidatorException extends Exception {

    private static final long serialVersionUID = 1L;
	private final ErrorCode code;
    private final Object[] args;

    public LoadValidatorException(ErrorCode code) {
        super();
        this.code = code;
        this.args = null;
    }

    public LoadValidatorException(ErrorCode code, Object... args) {
        super();
        this.code = code;
        this.args = args;
    }

    public ErrorCode getCode() {
        return code;
    }

    public Object[] getArgs() {
        return args;
    }

    public enum ErrorCode {

        VALUE_NUMBER_FORMAT_EXCEPTION("El valor del campo [{0}] debe ser numerico."),
        VALUE_RUT_FORMAT_EXCEPTION("El valor del campo [{0}] no tiene un rut válido.\nEjemplo: 99999999-9"),
        VALUE_REQUIRED("El valor del campo [{0}] es requerido"),
        VALUE_VIOLATES_REGEX("El valor del campo [{0}] no cumple con los valores permitidos."),
        VALUE_VIOLATES_MAXLENGTH("El valor del campo [{0}] excede el máximo permitido de {1} caracteres."),
        RECORD_DOES_NOT_EXIST_IN_DATABASE("El valor del campo [{0}] no existe en la base de datos."),
        VALUE_DATE_FORMAT_EXCEPTION("El valor del campo [{0}] no tiene un formato de fecha valido {1}\nEjemplo: 11-ENERO-2016"),
        RECORD_DUPLICATED("El registro {0} se encuentra duplicado en la hoja");

        private final String message;

        ErrorCode(String message) {
            this.message = message;
        }

        public String getMessage() {
            return message;
        }

        public String getMessage(Object... args) {
            StringBuilder sb = new StringBuilder();
            sb.append(message);
            if (args != null) {
                for (int i = 0; i < args.length; i++) {
                    Utils.replaceAll(sb, "{" + i + "}", String.valueOf(args[i]));
                }
            }
            return sb.toString();
        }
    }
}
