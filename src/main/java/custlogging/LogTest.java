package custlogging;

import java.util.Date;

public class LogTest {

    private static final Log log = LogFactory.getLogger();

    /**
     * @param args
     */
    public static void main(String[] args) {
        log.info("test started {0}", new Date());
        try {
            throw new Exception("Testing exception");
        } catch (Exception e) {
            log.error("exception caught", e);
        }
    }
}
