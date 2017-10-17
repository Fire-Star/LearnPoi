package cn.domarvel.exception;

public class BaseException extends Exception{
    private String errorType;//这是错误类型，一般是前端会用的一个东西，它标志着 key 。
    private String errorMessage;//这是错误信息，代表具体错误解释。

    public BaseException(String errorType, String errorMessage) {
        this.errorType = errorType;
        this.errorMessage = errorMessage;
    }

    public BaseException() {
    }

    public BaseException(String message) {
        super(message);
    }

    public BaseException(String message, Throwable cause) {
        super(message, cause);
    }

    public BaseException(Throwable cause) {
        super(cause);
    }
}
