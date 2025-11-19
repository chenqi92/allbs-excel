package cn.allbs.excel.util;

import lombok.extern.slf4j.Slf4j;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.function.Function;

/**
 * Class Metadata Cache
 * <p>
 * Provides caching for class metadata to improve performance by avoiding repeated reflection analysis.
 * Thread-safe implementation using ConcurrentHashMap.
 * </p>
 *
 * @author ChenQi
 * @since 2025-11-19
 */
@Slf4j
public class ClassMetadataCache {

	/**
	 * Cache for different types of metadata
	 * Key: Cache type name (e.g., "DynamicHeader", "ListExpand")
	 * Value: Map of Class -> Metadata
	 */
	private static final Map<String, Map<Class<?>, Object>> CACHE_REGISTRY = new ConcurrentHashMap<>();

	/**
	 * Cache statistics
	 */
	private static final Map<String, CacheStats> STATS_REGISTRY = new ConcurrentHashMap<>();

	/**
	 * Get or compute metadata from cache
	 *
	 * @param cacheType  Cache type identifier
	 * @param clazz      Class to analyze
	 * @param computer   Function to compute metadata if not cached
	 * @param <T>        Metadata type
	 * @return Cached or computed metadata
	 */
	@SuppressWarnings("unchecked")
	public static <T> T getOrCompute(String cacheType, Class<?> clazz, Function<Class<?>, T> computer) {
		// Get or create cache for this type
		Map<Class<?>, Object> cache = CACHE_REGISTRY.computeIfAbsent(cacheType, k -> new ConcurrentHashMap<>());

		// Get or create stats
		CacheStats stats = STATS_REGISTRY.computeIfAbsent(cacheType, k -> new CacheStats());

		// Try to get from cache
		Object cached = cache.get(clazz);
		if (cached != null) {
			stats.recordHit();
			log.trace("Cache hit for {}: {}", cacheType, clazz.getName());
			return (T) cached;
		}

		// Cache miss - compute metadata
		stats.recordMiss();
		log.debug("Cache miss for {}: {}, computing metadata...", cacheType, clazz.getName());

		T metadata = computer.apply(clazz);
		cache.put(clazz, metadata);

		log.debug("Cached metadata for {}: {} (cache size: {})", cacheType, clazz.getName(), cache.size());
		return metadata;
	}

	/**
	 * Clear cache for specific type
	 *
	 * @param cacheType Cache type to clear
	 */
	public static void clear(String cacheType) {
		Map<Class<?>, Object> cache = CACHE_REGISTRY.get(cacheType);
		if (cache != null) {
			int size = cache.size();
			cache.clear();
			log.info("Cleared cache for {}: {} entries removed", cacheType, size);
		}
	}

	/**
	 * Clear all caches
	 */
	public static void clearAll() {
		int totalEntries = CACHE_REGISTRY.values().stream().mapToInt(Map::size).sum();
		CACHE_REGISTRY.clear();
		STATS_REGISTRY.clear();
		log.info("Cleared all caches: {} entries removed", totalEntries);
	}

	/**
	 * Get cache statistics
	 *
	 * @param cacheType Cache type
	 * @return Cache statistics
	 */
	public static CacheStats getStats(String cacheType) {
		return STATS_REGISTRY.getOrDefault(cacheType, new CacheStats());
	}

	/**
	 * Get cache size for specific type
	 *
	 * @param cacheType Cache type
	 * @return Number of cached entries
	 */
	public static int getCacheSize(String cacheType) {
		Map<Class<?>, Object> cache = CACHE_REGISTRY.get(cacheType);
		return cache != null ? cache.size() : 0;
	}

	/**
	 * Print cache statistics
	 */
	public static void printStats() {
		log.info("=== Cache Statistics ===");
		STATS_REGISTRY.forEach((type, stats) -> {
			int size = getCacheSize(type);
			double hitRate = stats.getHitRate();
			log.info("{}: size={}, hits={}, misses={}, hit_rate={:.2f}%", type, size, stats.hits, stats.misses,
					hitRate * 100);
		});
	}

	/**
	 * Cache statistics
	 */
	public static class CacheStats {

		private long hits = 0;

		private long misses = 0;

		void recordHit() {
			hits++;
		}

		void recordMiss() {
			misses++;
		}

		public long getHits() {
			return hits;
		}

		public long getMisses() {
			return misses;
		}

		public double getHitRate() {
			long total = hits + misses;
			return total == 0 ? 0.0 : (double) hits / total;
		}

	}

	/**
	 * Cache type constants
	 */
	public static class CacheType {

		public static final String DYNAMIC_HEADER = "DynamicHeader";

		public static final String LIST_EXPAND = "ListExpand";

		public static final String MULTI_SHEET_RELATION = "MultiSheetRelation";

		public static final String FLATTEN_LIST = "FlattenList";

		public static final String FLATTEN_FIELD = "FlattenField";

	}

}
