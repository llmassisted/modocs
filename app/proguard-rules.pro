# MoDocs ProGuard Rules
# Add project specific ProGuard rules here.

# Keep Hilt generated code
-keep class dagger.hilt.** { *; }
-keep class javax.inject.** { *; }
-keep class * extends dagger.hilt.android.internal.managers.ViewComponentManager$FragmentContextWrapper { *; }

# PDFBox Android — optional JPEG2000 codec not used
-dontwarn com.gemalto.jp2.JP2Decoder
-dontwarn com.gemalto.jp2.JP2Encoder

# Apache POI — log4j excluded at dependency level, suppress R8 warnings
-dontwarn org.apache.logging.log4j.**
