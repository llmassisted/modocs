plugins {
    alias(libs.plugins.android.library)
    alias(libs.plugins.kotlin.android)
}

android {
    namespace = "com.modocs.core.common"
    compileSdk = 35

    defaultConfig {
        minSdk = 26
    }

    compileOptions {
        sourceCompatibility = JavaVersion.VERSION_17
        targetCompatibility = JavaVersion.VERSION_17
    }

    kotlinOptions {
        jvmTarget = "17"
    }
}

dependencies {
    implementation(libs.pdfbox.android)
    // FileProvider, for sharing documents out to other apps
    implementation(libs.androidx.core.ktx)

    implementation(libs.kotlinx.coroutines.core)
    implementation(libs.kotlinx.coroutines.android)

    // OOXML decryption (password-protected DOCX/XLSX/PPTX)
    //
    // Only log4j-core is excluded, never the whole group. POI 5.x compiles
    // against the log4j-api facade and initialises a static LogManager logger in
    // POIFSFileSystem, so dropping the API leaves a dangling reference that
    // throws NoClassDefFoundError the first time any OOXML file is opened.
    // log4j-core is the heavyweight implementation (and the Log4Shell surface);
    // without it the API no-ops, which is what we want on Android.
    implementation(libs.poi) {
        exclude(group = "org.apache.logging.log4j", module = "log4j-core")
        exclude(group = "commons-logging")
    }
    implementation(libs.poi.ooxml.lite) {
        exclude(group = "org.apache.logging.log4j", module = "log4j-core")
        exclude(group = "commons-logging")
    }
}
