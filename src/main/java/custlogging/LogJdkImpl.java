package custlogging;

import java.util.logging.Level;
import java.util.logging.Logger;

class LogJdkImpl implements Log {

    /**
     * Level-mapping for debug logging
     */
    static final Level DEBUG = Level.FINE;
    /**
     * Level-mapping for info logging
     */
    static final Level INFO = Level.INFO;
    /**
     * Level-mapping for config logging
     */
    static final Level CONFIG = Level.CONFIG;
    /**
     * Level-mapping for warn logging
     */
    static final Level WARN = Level.WARNING;
    /**
     * Level-mapping for error logging
     */
    static final Level ERROR = Level.SEVERE;

    final Logger logger;

    private LogJdkImpl(String name) {
        this.logger = Logger.getLogger(name);
    }

    static LogJdkImpl create(String name) {
        return new LogJdkImpl(name);
    }

    @Override
    public void debug(String msg, Object... args) {
        log(DEBUG, msg, args);
    }

    @Override
    public void debug(String msg, Throwable exc) {
        log(DEBUG, msg, exc);
    }

    @Override
    public boolean isDebugEnabled() {
        return logger.isLoggable(DEBUG);
    }

    @Override
    public void info(String msg, Object... args) {
        log(INFO, msg, args);
    }

    @Override
    public void info(String msg, Throwable exc) {
        log(INFO, msg, exc);
    }

    @Override
    public boolean isInfoEnabled() {
        return logger.isLoggable(INFO);
    }

    @Override
    public void config(String msg, Object... args) {
        log(CONFIG, msg, args);
    }

    @Override
    public void config(String msg, Throwable exc) {
        log(CONFIG, msg, exc);
    }

    @Override
    public boolean isConfigEnabled() {
        return logger.isLoggable(CONFIG);
    }

    @Override
    public void warn(String msg, Object... args) {
        log(WARN, msg, args);
    }

    @Override
    public void warn(String msg, Throwable exc) {
        log(WARN, msg, exc);
    }

    @Override
    public boolean isWarnEnabled() {
        return logger.isLoggable(WARN);
    }

    @Override
    public void error(String msg, Object... args) {
        log(ERROR, msg, args);
    }

    @Override
    public void error(String msg, Throwable exc) {
        log(ERROR, msg, exc);
    }

    @Override
    public boolean isErrorEnabled() {
        return logger.isLoggable(ERROR);
    }

    private void log(Level level, String msg, Object... args) {
        if (logger.isLoggable(level)) {
            StackTraceElement ste = Thread.currentThread().getStackTrace()[3];
            String methodName = ste.getMethodName() + (ste.getLineNumber() > 0 ? ":" + ste.getLineNumber() : "");
            logger.logp(level, ste.getClassName(), methodName, msg, args);
        }
    }

    private void log(Level level, String msg, Throwable exc) {
        if (logger.isLoggable(level)) {
            StackTraceElement ste = Thread.currentThread().getStackTrace()[3];
            String methodName = ste.getMethodName() + (ste.getLineNumber() > 0 ? ":" + ste.getLineNumber() : "");
            logger.logp(level, ste.getClassName(), methodName, msg, exc);
        }
    }
}
