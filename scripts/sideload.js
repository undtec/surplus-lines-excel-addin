#!/usr/bin/env node
/**
 * Sideload script for macOS Excel Add-in
 * Copies manifest to Excel's wef directory (works around cross-device link errors)
 */

const fs = require('fs');
const path = require('path');
const os = require('os');

// Use dev manifest for local development
const MANIFEST_FILE = process.argv[2] || 'manifest.dev.xml';
const ADDIN_ID = 'a1b2c3d4-e5f6-7890-abcd-ef1234567890';

// Get the wef directory based on platform
function getWefDirectory() {
    const platform = os.platform();
    const homeDir = os.homedir();

    if (platform === 'darwin') {
        // macOS - Excel from App Store (sandboxed)
        const sandboxedPath = path.join(
            homeDir,
            'Library/Containers/com.microsoft.Excel/Data/Documents/wef'
        );

        // macOS - Excel from Office installer (non-sandboxed)
        const nonSandboxedPath = path.join(
            homeDir,
            'Library/Containers/com.microsoft.Excel/Data/Library/Application Support/Microsoft/Office/16.0/Wef'
        );

        // Try sandboxed path first (App Store version)
        if (fs.existsSync(path.dirname(sandboxedPath))) {
            return sandboxedPath;
        }

        // Fall back to non-sandboxed path
        return nonSandboxedPath;
    } else if (platform === 'win32') {
        // Windows
        return path.join(
            process.env.LOCALAPPDATA || path.join(homeDir, 'AppData/Local'),
            'Microsoft/Office/16.0/Wef'
        );
    }

    throw new Error(`Unsupported platform: ${platform}`);
}

function sideload() {
    const manifestPath = path.join(process.cwd(), MANIFEST_FILE);

    // Check if manifest exists
    if (!fs.existsSync(manifestPath)) {
        console.error(`Error: ${MANIFEST_FILE} not found in current directory`);
        process.exit(1);
    }

    const wefDir = getWefDirectory();
    const destPath = path.join(wefDir, `${ADDIN_ID}.manifest.xml`);

    console.log('Sideloading Office Add-in...');
    console.log(`  Source: ${manifestPath}`);
    console.log(`  Destination: ${destPath}`);

    // Create wef directory if it doesn't exist
    if (!fs.existsSync(wefDir)) {
        console.log(`  Creating directory: ${wefDir}`);
        fs.mkdirSync(wefDir, { recursive: true });
    }

    // Copy manifest (not link) to avoid cross-device link errors
    try {
        fs.copyFileSync(manifestPath, destPath);
        console.log('✓ Add-in manifest installed successfully!');
        console.log('');
        console.log('Next steps:');
        console.log('  1. Start Excel');
        console.log('  2. Go to Insert > Add-ins > My Add-ins');
        console.log('  3. Look for "Surplus Lines Tax Calculator" in Shared Folder');
        console.log('  4. If not visible, restart Excel and try again');
        console.log('');
        console.log('Make sure the dev server is running: npm run dev-server');
    } catch (error) {
        console.error(`Error copying manifest: ${error.message}`);
        process.exit(1);
    }
}

sideload();
