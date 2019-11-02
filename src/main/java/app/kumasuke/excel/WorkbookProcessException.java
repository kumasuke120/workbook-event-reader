package app.kumasuke.excel;

/**
 * An exception denotes errors happened during the reading process of a {@link WorkbookEventReader}
 */
@SuppressWarnings("WeakerAccess")
public class WorkbookProcessException extends WorkbookEventReaderException {

    WorkbookProcessException(Throwable cause) {
        super(cause);
    }

}
