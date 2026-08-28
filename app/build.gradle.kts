plugins {
    alias(libs.plugins.android.application)
    alias(libs.plugins.kotlin.android)
    alias(libs.plugins.kotlin.compose)
    alias(libs.plugins.ksp)
    alias(libs.plugins.hilt)
}

val releaseStoreFilePath = providers.gradleProperty("MODOCS_RELEASE_STORE_FILE").orNull
    ?: System.getenv("MODOCS_RELEASE_STORE_FILE")
val releaseStorePassword = providers.gradleProperty("MODOCS_RELEASE_STORE_PASSWORD").orNull
    ?: System.getenv("MODOCS_RELEASE_STORE_PASSWORD")
val releaseKeyAlias = providers.gradleProperty("MODOCS_RELEASE_KEY_ALIAS").orNull
    ?: System.getenv("MODOCS_RELEASE_KEY_ALIAS")
val releaseKeyPassword = providers.gradleProperty("MODOCS_RELEASE_KEY_PASSWORD").orNull
    ?: System.getenv("MODOCS_RELEASE_KEY_PASSWORD")
val hasReleaseSigning = listOf(
    releaseStoreFilePath,
    releaseStorePassword,
    releaseKeyAlias,
    releaseKeyPassword,
).all { !it.isNullOrBlank() }

android {
    namespace = "com.modocs.app"
    compileSdk = 35

    defaultConfig {
        applicationId = "com.modocs.app"
        minSdk = 26
        targetSdk = 35
        versionCode = 174
        versionName = "1.74"

        testInstrumentationRunner = "androidx.test.runner.AndroidJUnitRunner"
    }

    signingConfigs {
        create("release") {
            if (hasReleaseSigning) {
                storeFile = file(releaseStoreFilePath!!)
                storePassword = releaseStorePassword
                keyAlias = releaseKeyAlias
                keyPassword = releaseKeyPassword
            }
        }
    }

    buildTypes {
        release {
            isMinifyEnabled = true
            if (hasReleaseSigning) {
                signingConfig = signingConfigs.getByName("release")
            }
            proguardFiles(
                getDefaultProguardFile("proguard-android-optimize.txt"),
                "proguard-rules.pro"
            )
        }
    }

    compileOptions {
        sourceCompatibility = JavaVersion.VERSION_17
        targetCompatibility = JavaVersion.VERSION_17
    }

    kotlinOptions {
        jvmTarget = "17"
    }

    buildFeatures {
        compose = true
    }
}

dependencies {
    implementation(project(":core:core-common"))
    implementation(project(":core:core-ui"))
    implementation(project(":core:core-model"))
    implementation(project(":core:core-storage"))
    implementation(project(":feature:feature-home"))
    implementation(project(":feature:feature-pdf"))
    implementation(project(":feature:feature-docx"))
    implementation(project(":feature:feature-xlsx"))
    implementation(project(":feature:feature-pptx"))
    implementation(project(":fonts"))

    implementation(libs.androidx.core.ktx)
    implementation(libs.androidx.appcompat)
    implementation(libs.androidx.activity.compose)

    implementation(platform(libs.compose.bom))
    implementation(libs.compose.ui)
    implementation(libs.compose.ui.graphics)
    implementation(libs.compose.ui.tooling.preview)
    implementation(libs.compose.foundation)
    implementation(libs.compose.material3)
    implementation(libs.compose.material.icons.extended)
    implementation(libs.compose.material3.adaptive)
    implementation(libs.compose.material3.adaptive.layout)
    implementation(libs.compose.material3.adaptive.navigation)
    implementation(libs.compose.material3.adaptive.navigation.suite)
    implementation(libs.compose.material3.window.size)

    implementation(libs.lifecycle.runtime.ktx)
    implementation(libs.lifecycle.runtime.compose)
    implementation(libs.lifecycle.viewmodel.compose)
    implementation(libs.navigation.compose)

    implementation(libs.hilt.android)
    ksp(libs.hilt.compiler)
    implementation(libs.hilt.navigation.compose)

    implementation(libs.kotlinx.coroutines.core)
    implementation(libs.kotlinx.coroutines.android)

    debugImplementation(libs.compose.ui.tooling)

    testImplementation(libs.junit)
}
