package com.github.kumasuke120.excel.handler;

import org.jetbrains.annotations.ApiStatus;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@ApiStatus.Experimental
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface WorkbookRecord {

    int startSheet() default 0;

    int endSheet() default 255;

    int startRow() default 0;

    int endRow() default Integer.MAX_VALUE;

    int startColumn() default 0;

    int endColumn() default 16384;

    @Target(ElementType.FIELD)
    @Retention(RetentionPolicy.RUNTIME)
    @interface Metadata {

        MetadataType value();

    }

    @Target(ElementType.FIELD)
    @Retention(RetentionPolicy.RUNTIME)
    @interface Property {

        int column();

        boolean strict() default false;

        CellValueType valueType() default CellValueType.AUTO;

        String valueMethod() default "";

    }

    enum CellValueType {

        AUTO,

        BOOLEAN,

        INTEGER,

        LONG,

        DOUBLE,

        DECIMAL,

        STRING,

        RAW_STRING,

        TIME,

        DATE,

        DATETIME

    }

    enum MetadataType {

        SHEET_INDEX,

        SHEET_NAME,

        ROW_NUMBER
    }

}
