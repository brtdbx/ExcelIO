package custlogging;

public interface Log {
    void debug(String msg, Object... args);
    void debug(String msg, Throwable exc);
    boolean isDebugEnabled();
 
    void info(String msg, Object... args);
    void info(String msg, Throwable exc);
    boolean isInfoEnabled();
 
    void config(String msg, Object... args);
    void config(String msg, Throwable exc);
    boolean isConfigEnabled();
 
    void warn(String msg, Object... args);
    void warn(String msg, Throwable exc);
    boolean isWarnEnabled();
 
    void error(String msg, Object... args);
    void error(String msg, Throwable exc);
    boolean isErrorEnabled();
}