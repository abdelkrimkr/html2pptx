const path = require('path');

/**
 * Security utilities for path validation
 */

/**
 * Validates a path for security risks like null bytes
 * @param {string} filePath - The path to validate
 * @throws {Error} If the path is invalid or contains malicious characters
 */
function validatePath(filePath) {
    if (typeof filePath !== 'string') {
        throw new Error('Path must be a string');
    }

    // Check for null bytes which can be used to bypass path validation
    if (filePath.includes('\0')) {
        throw new Error('Path contains null bytes');
    }

    return true;
}

/**
 * Checks if a path is safe and within an optional base directory
 * @param {string} filePath - The path to check
 * @param {string} baseDir - Optional base directory to restrict access to
 * @returns {boolean} True if the path is safe
 */
function isPathSafe(filePath, baseDir) {
    if (!baseDir) return true;

    const resolvedPath = path.resolve(filePath);
    const resolvedBase = path.resolve(baseDir);

    const relative = path.relative(resolvedBase, resolvedPath);

    // Path is safe if it doesn't start with '..' and is not absolute
    // (relative paths that stay within the base directory)
    return !relative.startsWith('..') && !path.isAbsolute(relative);
}

module.exports = {
    validatePath,
    isPathSafe
};
