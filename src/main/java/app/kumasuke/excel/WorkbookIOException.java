package app.kumasuke.excel;

/**
 * An exception which related to IO when a {@link WorkbookEventReader} reading a workbook
 */
@SuppressWarnings("WeakerAccess")
public class WorkbookIOException extends WorkbookEventReaderException {
    WorkbookIOException(String message, Throwable cause) {
        super(message, cause);
    }
}
