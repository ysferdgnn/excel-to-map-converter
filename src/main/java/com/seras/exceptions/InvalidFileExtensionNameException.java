package com.seras.exceptions;

public class InvalidFileExtensionNameException extends  Exception{
    public InvalidFileExtensionNameException() {
        super();
    }

    public InvalidFileExtensionNameException(String message) {
        super(message);
    }
}
