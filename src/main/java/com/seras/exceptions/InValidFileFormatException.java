package com.seras.exceptions;

public class InValidFileFormatException extends Exception{

    public InValidFileFormatException(String message) {
        super(message);
    }

    public InValidFileFormatException() {
        super();
    }
}
