package app.kumasuke.excel;

/**
 * An exception occurred when the {@link WorkbookEventReader} stays in a particular state where some kind of
 * operation cannot be performed
 */
@SuppressWarnings("WeakerAccess")
public class IllegalReaderStateException extends WorkbookEventReaderException {

    IllegalReaderStateException(String message) {
        super(message);
    }

}
