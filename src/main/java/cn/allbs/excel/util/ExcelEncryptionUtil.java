package cn.allbs.excel.util;

import cn.allbs.excel.annotation.ExcelEncryption;
import cn.allbs.excel.exception.ExcelExportException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * Excel Encryption Utility
 * <p>
 * Provides Excel file encryption functionality
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
@Slf4j
public class ExcelEncryptionUtil {

	/**
	 * Encrypt Excel file
	 *
	 * @param sourceFile     Source Excel file
	 * @param password       Encryption password
	 * @param algorithm      Encryption algorithm
	 * @return Encrypted file
	 */
	public static File encryptFile(File sourceFile, String password, ExcelEncryption.EncryptionAlgorithm algorithm) {
		try {
			File encryptedFile = File.createTempFile("encrypted_", ".xlsx");
			encryptedFile.deleteOnExit();

			encryptFile(sourceFile, encryptedFile, password, algorithm);

			return encryptedFile;
		} catch (Exception e) {
			log.error("Failed to encrypt Excel file: {}", sourceFile.getAbsolutePath(), e);
			throw ExcelExportException.builder()
				.message("Excel file encryption failed: " + e.getMessage())
				.cause(e)
				.build();
		}
	}

	/**
	 * Encrypt Excel file
	 *
	 * @param sourceFile     Source Excel file
	 * @param targetFile     Target encrypted file
	 * @param password       Encryption password
	 * @param algorithm      Encryption algorithm
	 */
	public static void encryptFile(File sourceFile, File targetFile, String password,
	                               ExcelEncryption.EncryptionAlgorithm algorithm) {
		try (FileInputStream fis = new FileInputStream(sourceFile);
		     FileOutputStream fos = new FileOutputStream(targetFile)) {

			encryptStream(fis, fos, password, algorithm);

			log.info("Successfully encrypted Excel file: {} -> {}", sourceFile.getName(), targetFile.getName());
		} catch (Exception e) {
			log.error("Failed to encrypt Excel file: {} -> {}", sourceFile.getAbsolutePath(),
				targetFile.getAbsolutePath(), e);
			throw ExcelExportException.builder()
				.message("Excel file encryption failed: " + e.getMessage())
				.cause(e)
				.build();
		}
	}

	/**
	 * Encrypt Excel stream
	 *
	 * @param inputStream    Input stream
	 * @param outputStream   Output stream
	 * @param password       Encryption password
	 * @param algorithm      Encryption algorithm
	 */
	public static void encryptStream(InputStream inputStream, OutputStream outputStream, String password,
	                                 ExcelEncryption.EncryptionAlgorithm algorithm) {
		try {
			// Create POIFSFileSystem
			POIFSFileSystem fs = new POIFSFileSystem();

			// Select encryption mode
			EncryptionMode mode = algorithm == ExcelEncryption.EncryptionAlgorithm.AGILE
				? EncryptionMode.agile
				: EncryptionMode.standard;

			// Create encryption info
			EncryptionInfo info = new EncryptionInfo(mode);
			Encryptor enc = info.getEncryptor();
			enc.confirmPassword(password);

			// Load source file
			try (OPCPackage opc = OPCPackage.open(inputStream)) {
				// Encrypt to POIFSFileSystem
				try (OutputStream os = enc.getDataStream(fs)) {
					opc.save(os);
				}
			}

			// Write encrypted file system to output stream
			fs.writeFilesystem(outputStream);

			log.debug("Successfully encrypted Excel stream using {} algorithm", algorithm);
		} catch (Exception e) {
			log.error("Failed to encrypt Excel stream", e);
			throw ExcelExportException.builder()
				.message("Excel stream encryption failed: " + e.getMessage())
				.cause(e)
				.build();
		}
	}

	/**
	 * Decrypt Excel file
	 *
	 * @param encryptedFile  Encrypted Excel file
	 * @param password       Decryption password
	 * @return Decrypted file
	 */
	public static File decryptFile(File encryptedFile, String password) {
		try {
			File decryptedFile = File.createTempFile("decrypted_", ".xlsx");
			decryptedFile.deleteOnExit();

			decryptFile(encryptedFile, decryptedFile, password);

			return decryptedFile;
		} catch (Exception e) {
			log.error("Failed to decrypt Excel file: {}", encryptedFile.getAbsolutePath(), e);
			throw new RuntimeException("Excel file decryption failed", e);
		}
	}

	/**
	 * Decrypt Excel file
	 *
	 * @param encryptedFile  Encrypted Excel file
	 * @param targetFile     Target decrypted file
	 * @param password       Decryption password
	 */
	public static void decryptFile(File encryptedFile, File targetFile, String password) {
		try (POIFSFileSystem fs = new POIFSFileSystem(encryptedFile, true)) {
			EncryptionInfo info = new EncryptionInfo(fs);
			org.apache.poi.poifs.crypt.Decryptor d = org.apache.poi.poifs.crypt.Decryptor.getInstance(info);

			if (!d.verifyPassword(password)) {
				throw new IllegalArgumentException("Invalid password");
			}

			try (InputStream dataStream = d.getDataStream(fs);
			     OPCPackage opc = OPCPackage.open(dataStream);
			     FileOutputStream fos = new FileOutputStream(targetFile)) {

				opc.save(fos);
			}

			log.info("Successfully decrypted Excel file: {} -> {}", encryptedFile.getName(), targetFile.getName());
		} catch (Exception e) {
			log.error("Failed to decrypt Excel file: {} -> {}", encryptedFile.getAbsolutePath(),
				targetFile.getAbsolutePath(), e);
			throw new RuntimeException("Excel file decryption failed", e);
		}
	}

	/**
	 * Check if Excel file is encrypted
	 *
	 * @param file Excel file
	 * @return true if encrypted, false otherwise
	 */
	public static boolean isEncrypted(File file) {
		try (POIFSFileSystem fs = new POIFSFileSystem(file, true)) {
			return fs.getRoot().hasEntry(org.apache.poi.poifs.crypt.Decryptor.DEFAULT_POIFS_ENTRY);
		} catch (Exception e) {
			// If cannot open as POIFS, it's not encrypted
			return false;
		}
	}
}
