/** Custom utility scripts bundled as modules */

/**
 * PropertyCache module.
 * 
 * A utility module to manage and store persistent state.
 * Leverages both PropertiesService and CacheService.
 * 
 * Caches properties to get around PropertyService 
 * read quotas limitations.
 *
 * All property values are processed as JSON serialized strings.
 *
 */
 
(function(context) {
    
    /**
     * Constructor
     */ 
    context.PropertyCache = function() {
        this.properties = PropertiesService.getDocumentProperties() || PropertiesService.getUserProperties();
        this.cache = CacheService.getDocumentCache() || CacheService.getUserCache();
    };
    
    /**
     * Retrieves a property from the cache.
     *
     * @param {String} key Key used to retrieve the property.
     */
    context.PropertyCache.prototype.get = function(key) {
        
        // check cache
        var cached = this.cache.get(key);
        
        if (!cached) {
            cached = this.properties.getProperty(key);
            cached && this.cache.put(key, cached, SIX_HOURS_IN_SECONDS);
        }
        
        return JSON.parse(cached);
    };
    
    /**
     * Puts a property into the cache.
     *
     * @param {String}  key                 - Key used to name and store the property
     * @param {String}  value               - Value of the propety
     * @param {Boolean} addToUserProperties - If true cached values are mirrored in properties object.
     */
    context.PropertyCache.prototype.put = function(key, value, addToUserProperties) {
        
        var v = JSON.stringify(value);
        
        this.cache.remove(key);
        this.cache.put(key, v, SIX_HOURS_IN_SECONDS);
        
        if (addToUserProperties) {
            this.properties.setProperty(key, v);
        }
        
    };
    
    /**
     * Removes a property from the cache
     *
     * @param {String}  key                      - Key used to reference a stored property
     * @param {Boolean} removeFromUserProperties - If true cached values are removed from properties object.
     */
    context.PropertyCache.prototype.remove = function(key, removeFromUserProperties) {
        
        this.cache.remove(key);
        
        if (removeFromUserProperties) {
            this.properties.deleteProperty(key);
        }
    };
    
})(this);