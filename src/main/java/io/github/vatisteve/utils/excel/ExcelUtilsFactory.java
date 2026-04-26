package io.github.vatisteve.utils.excel;

import io.github.vatisteve.utils.excel.loader.ExcelLoader;
import io.github.vatisteve.utils.excel.loader.ExcelLoaderImpl;
import io.github.vatisteve.utils.excel.writer.ExcelWriter;
import io.github.vatisteve.utils.excel.writer.ExcelWriterConfiguration;
import io.github.vatisteve.utils.excel.writer.ExcelWriterException;
import io.github.vatisteve.utils.excel.writer.ExcelWriterImpl;
import org.apache.poi.EncryptedDocumentException;

import java.io.IOException;
import java.io.InputStream;

/**
 * Factory class for creating instances of Excel-related utilities such as ExcelLoader and ExcelWriter.
 * This class provides a centralized way to create these objects with default or custom configurations.
 * This class cannot be instantiated.
 *
 * @since May 23, 2023
 */
public final class ExcelUtilsFactory {

    private ExcelUtilsFactory() {
    }

    /**
     * Creates and returns an instance of {@code ExcelLoader} initialized with the given input stream.
     * The returned {@code ExcelLoader} object can be used to load and process Excel files.
     *
     * @param inputStream the {@code InputStream} from which the Excel document will be read.
     *                    This stream must point to a valid Excel file.
     * @return an instance of {@code ExcelLoader} for interacting with the Excel file.
     * @throws EncryptedDocumentException if the Excel file is encrypted and cannot be processed.
     * @throws IOException                if an I/O error occurs while reading the stream.
     */
    public static ExcelLoader createExcelLoader(InputStream inputStream) throws EncryptedDocumentException, IOException {
        return new ExcelLoaderImpl(inputStream);
    }

    /**
     * Creates and returns a new instance of {@code ExcelWriter} using the default configuration.
     * The returned {@code ExcelWriter} allows the creation and manipulation of Excel files,
     * such as writing data to cells, applying styles, and generating output in Excel format.
     *
     * @return a new instance of {@code ExcelWriter} with the default configuration.
     */
    public static ExcelWriter createExcelWriter() {
        return new ExcelWriterImpl(new ExcelWriterConfiguration.DefaultConfiguration());
    }

    /**
     * Creates and returns a new instance of {@code ExcelWriter} using the provided configuration.
     * The returned {@code ExcelWriter} allows for creating and manipulating Excel files
     * based on the specified configuration, such as writing data to cells and applying styles.
     *
     * @param configuration the configuration to initialize the {@code ExcelWriter} instance.
     *                      This contains settings like default styles, formatting rules,
     *                      and other customization options.
     * @return a new instance of {@code ExcelWriter} configured based on the given {@code ExcelWriterConfiguration}.
     */
    public static ExcelWriter createExcelWriter(ExcelWriterConfiguration configuration) {
        return new ExcelWriterImpl(configuration);
    }

    /**
     * Creates and returns a new instance of {@code ExcelWriter} initialized with the given input stream
     * and using the default configuration.
     * The returned {@code ExcelWriter} can be used for writing data to an existing or new Excel file.
     *
     * @param inputStream the {@code InputStream} containing the Excel file to be written or modified.
     *                    The stream must point to a valid Excel file.
     * @return a new instance of {@code ExcelWriter} initialized with the specified input stream
     * and default configuration.
     * @throws ExcelWriterException if an error occurs while creating the {@code ExcelWriter},
     *                              such as invalid input stream data or an I/O issue.
     */
    public static ExcelWriter createExcelWriter(InputStream inputStream) throws ExcelWriterException {
        return new ExcelWriterImpl(inputStream, new ExcelWriterConfiguration.DefaultConfiguration());
    }

    /**
     * Creates and returns a new instance of {@code ExcelWriter} initialized with the given input stream
     * and configuration. The returned {@code ExcelWriter} allows for writing and modifying Excel files
     * based on the specified configuration, such as applying custom styles or formatting rules.
     *
     * @param inputStream   the {@code InputStream} containing the Excel file to be written or modified.
     *                      The input stream must point to a valid Excel file.
     * @param configuration the {@code ExcelWriterConfiguration} to initialize the {@code ExcelWriter}.
     *                      Contains settings like default styles, formatting rules,
     *                      or specific writer configurations.
     * @return a new instance of {@code ExcelWriter} configured with the provided input stream
     * and custom configuration.
     * @throws ExcelWriterException if an error occurs while creating the {@code ExcelWriter},
     *                              such as invalid input stream data, configuration issues, or I/O errors.
     */
    public static ExcelWriter createExcelWriter(InputStream inputStream, ExcelWriterConfiguration configuration) throws ExcelWriterException {
        return new ExcelWriterImpl(inputStream, configuration);
    }
}
