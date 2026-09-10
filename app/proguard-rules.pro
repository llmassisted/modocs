# MoDocs ProGuard Rules
# Add project specific ProGuard rules here.

# Keep Hilt generated code
-keep class dagger.hilt.** { *; }
-keep class javax.inject.** { *; }
-keep class * extends dagger.hilt.android.internal.managers.ViewComponentManager$FragmentContextWrapper { *; }

# PDFBox Android — optional JPEG2000 codec not used
-dontwarn com.gemalto.jp2.JP2Decoder
-dontwarn com.gemalto.jp2.JP2Encoder

# Apache POI logs through the log4j-api facade and initialises a static logger,
# so log4j-api must stay on the classpath — only log4j-core is excluded (see
# core-common/build.gradle.kts). Keep the API's entry points so R8 cannot strip
# what POI resolves at class-init time.
-keep class org.apache.logging.log4j.LogManager { *; }
-keep class org.apache.logging.log4j.Logger { *; }
-keep class org.apache.logging.log4j.spi.** { *; }

# log4j-core is deliberately absent; the API references it reflectively.
# This is scoped to .core on purpose: a blanket -dontwarn org.apache.logging.log4j.**
# previously hid the missing LogManager that broke every OOXML open in v1.74.
-dontwarn org.apache.logging.log4j.core.**

# log4j-api carries OSGi and bnd annotations for container environments that do
# not exist on Android; it guards the lookups and degrades gracefully. Listed by
# exact class rather than a wildcard so a genuinely missing class still fails the
# build instead of being silently swallowed.
-dontwarn aQute.bnd.annotation.spi.ServiceConsumer
-dontwarn aQute.bnd.annotation.spi.ServiceProvider
-dontwarn org.osgi.annotation.bundle.Export
-dontwarn org.osgi.annotation.versioning.Version
-dontwarn org.osgi.framework.Bundle
-dontwarn org.osgi.framework.BundleContext
-dontwarn org.osgi.framework.FrameworkUtil
-dontwarn org.osgi.framework.ServiceReference
