package io.github.vatisteve.utils.excel;

import java.io.IOException;
import java.io.InputStream;

import io.github.vatisteve.utils.excel.writer.ExcelWriter;
import io.github.vatisteve.utils.excel.writer.ExcelWriterConfiguration;
import io.github.vatisteve.utils.excel.writer.ExcelWriterException;
import io.github.vatisteve.utils.excel.writer.ExcelWriterImpl;
import org.apache.poi.EncryptedDocumentException;

import io.github.vatisteve.utils.excel.loader.ExcelLoader;
import io.github.vatisteve.utils.excel.loader.ExcelLoaderImpl;

/**
 * @author Steve
 * @since May 23, 2023
 *
 */
public final class ExcelUtilsFactory {

    private ExcelUtilsFactory() {}

    public static ExcelLoader createExcelLoader(InputStream inputStream) throws EncryptedDocumentException, IOException {
        return new ExcelLoaderImpl(inputStream);
    }

    public static ExcelWriter createExcelWriter() {
        return new ExcelWriterImpl(new ExcelWriterConfiguration.DefaultConfiguration());
    }

    public static ExcelWriter createExcelWriter(ExcelWriterConfiguration configuration) {
        return new ExcelWriterImpl(configuration);
    }

    public static ExcelWriter createExcelWriter(InputStream inputStream) throws ExcelWriterException {
        return new ExcelWriterImpl(inputStream, new ExcelWriterConfiguration.DefaultConfiguration());
    }

    public static ExcelWriter createExcelWriter(InputStream inputStream, ExcelWriterConfiguration configuration) throws ExcelWriterException {
        return new ExcelWriterImpl(inputStream, configuration);
    }
}
