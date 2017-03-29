package custlogging;

public class LogFactory {

    private LogFactory() {
    }

    /**
     * KISS factory method that determines logger name from the calling class
     * using stacktrace
     *
     * @return Log implementation for the calling class
     */
    public static Log getLogger() {
        StackTraceElement[] stackTrace = Thread.currentThread().getStackTrace();
        if (stackTrace.length > 2) {
            return getLogger(stackTrace[2].getClassName());
        } else {
            return getLogger(LogFactory.class);
        }
    }

    /**
     * Factory method for creating Log implementation for the class given
     *
     * @param clazz class to get Log implementation for
     * @return Log implementation for the class argument given
     */
    public static Log getLogger(Class<?> clazz) {
        return getLoggerImpl(clazz.getName());
    }

    /**
     * Factory method for creating Log implementation for the name given
     *
     * @param className class-name to get Log implementation for
     * @return Log implementation for the className argument given
     */
    public static Log getLogger(String className) {
        return getLoggerImpl(className);
    }

    /**
     * Internal factory method - this is where you change if you do not like the
     * default jdk-logging implementation
     *
     * @param name Name of logger
     * @return appropriate Log implementation for the name given
     */
    private static Log getLoggerImpl(String name) {
        return LogJdkImpl.create(name);
    }
}
