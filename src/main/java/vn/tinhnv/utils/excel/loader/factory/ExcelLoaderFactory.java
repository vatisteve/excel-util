package vn.tinhnv.utils.excel.loader.factory;

import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.EncryptedDocumentException;

import vn.tinhnv.utils.excel.loader.ExcelLoader;
import vn.tinhnv.utils.excel.loader.implementation.ExcelLoaderImpl;

/**
 * @author Steve
 * @since May 23, 2023
 *
 */
public final class ExcelLoaderFactory {

    private ExcelLoaderFactory() {}

    public static ExcelLoader createExcelLoader(InputStream inputStream) throws EncryptedDocumentException, IOException {
        return new ExcelLoaderImpl(inputStream);
    }
}
