package cn.allbs.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel Encryption Annotation
 * <p>
 * Encrypts the generated Excel file with a password
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
@Target({ElementType.METHOD, ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelEncryption {

	/**
	 * Encryption password
	 * <p>
	 * Supports SpEL expressions, e.g., "${excel.password}" or "#{T(java.util.UUID).randomUUID().toString()}"
	 * </p>
	 */
	String password();

	/**
	 * Whether encryption is enabled
	 * <p>
	 * Default: true
	 * </p>
	 */
	boolean enabled() default true;

	/**
	 * Encryption algorithm
	 * <p>
	 * Supported algorithms:
	 * - STANDARD: Standard Excel encryption (compatible with all Excel versions)
	 * - AGILE: Agile encryption (Excel 2010+, stronger security)
	 * </p>
	 */
	EncryptionAlgorithm algorithm() default EncryptionAlgorithm.STANDARD;

	/**
	 * Encryption algorithm type
	 */
	enum EncryptionAlgorithm {
		/**
		 * Standard Excel encryption (XOR obfuscation)
		 * Compatible with all Excel versions
		 */
		STANDARD,

		/**
		 * Agile encryption mode (AES encryption)
		 * Excel 2010+, stronger security
		 */
		AGILE
	}
}
