package com.seras.exceptions;

public class InvalidSpreadSheetFormatException extends Exception{
    public InvalidSpreadSheetFormatException() {
        super();
    }

    public InvalidSpreadSheetFormatException(String message) {
        super(message);
    }
}
