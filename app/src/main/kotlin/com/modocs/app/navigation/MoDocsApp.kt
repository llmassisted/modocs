package com.modocs.app.navigation

import android.app.Activity
import android.net.Uri
import androidx.activity.compose.BackHandler
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.padding
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.filled.Description
import androidx.compose.material.icons.filled.History
import androidx.compose.material.icons.filled.Settings
import androidx.compose.material.icons.outlined.Description
import androidx.compose.material.icons.outlined.History
import androidx.compose.material.icons.outlined.Settings
import androidx.compose.material3.Icon
import androidx.compose.material3.Scaffold
import androidx.compose.material3.Text
import androidx.compose.material3.adaptive.navigationsuite.NavigationSuiteScaffold
import androidx.compose.runtime.Composable
import androidx.compose.runtime.LaunchedEffect
import androidx.compose.runtime.getValue
import androidx.compose.runtime.mutableStateOf
import androidx.compose.runtime.saveable.rememberSaveable
import androidx.compose.runtime.setValue
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.vector.ImageVector
import androidx.compose.ui.platform.LocalContext
import androidx.navigation.NavType
import androidx.navigation.compose.NavHost
import androidx.navigation.compose.composable
import androidx.navigation.compose.rememberNavController
import androidx.navigation.navArgument
import com.modocs.core.common.DocumentType
import com.modocs.feature.docx.DocxViewerScreen
import com.modocs.feature.home.HomeScreen
import com.modocs.feature.home.SettingsScreen
import com.modocs.feature.pdf.PdfViewerScreen
import com.modocs.feature.pptx.PptxViewerScreen
import com.modocs.feature.xlsx.XlsxViewerScreen

object Routes {
    const val FILES = "files"
    const val RECENT = "recent"
    const val SETTINGS = "settings"
    const val PDF_VIEWER = "pdf_viewer?uri={uri}&name={name}"
    const val DOCX_VIEWER = "docx_viewer?uri={uri}&name={name}"
    const val XLSX_VIEWER = "xlsx_viewer?uri={uri}&name={name}"
    const val PPTX_VIEWER = "pptx_viewer?uri={uri}&name={name}"

    fun pdfViewer(uri: Uri, displayName: String? = null): String {
        val encoded = Uri.encode(uri.toString())
        val name = Uri.encode(displayName ?: "")
        return "pdf_viewer?uri=$encoded&name=$name"
    }

    fun docxViewer(uri: Uri, displayName: String? = null): String {
        val encoded = Uri.encode(uri.toString())
        val name = Uri.encode(displayName ?: "")
        return "docx_viewer?uri=$encoded&name=$name"
    }

    fun xlsxViewer(uri: Uri, displayName: String? = null): String {
        val encoded = Uri.encode(uri.toString())
        val name = Uri.encode(displayName ?: "")
        return "xlsx_viewer?uri=$encoded&name=$name"
    }

    fun pptxViewer(uri: Uri, displayName: String? = null): String {
        val encoded = Uri.encode(uri.toString())
        val name = Uri.encode(displayName ?: "")
        return "pptx_viewer?uri=$encoded&name=$name"
    }

    fun viewerForType(uri: Uri, documentType: DocumentType, displayName: String? = null): String {
        return when (documentType) {
            DocumentType.PDF -> pdfViewer(uri, displayName)
            DocumentType.DOCX -> docxViewer(uri, displayName)
            DocumentType.XLSX -> xlsxViewer(uri, displayName)
            DocumentType.PPTX -> pptxViewer(uri, displayName)
            else -> pdfViewer(uri, displayName) // fallback
        }
    }
}

enum class TopLevelDestination(
    val label: String,
    val selectedIcon: ImageVector,
    val unselectedIcon: ImageVector,
    val route: String,
) {
    FILES(
        label = "Files",
        selectedIcon = Icons.Filled.Description,
        unselectedIcon = Icons.Outlined.Description,
        route = Routes.FILES,
    ),
    RECENT(
        label = "Recent",
        selectedIcon = Icons.Filled.History,
        unselectedIcon = Icons.Outlined.History,
        route = Routes.RECENT,
    ),
    SETTINGS(
        label = "Settings",
        selectedIcon = Icons.Filled.Settings,
        unselectedIcon = Icons.Outlined.Settings,
        route = Routes.SETTINGS,
    ),
}

@Composable
fun MoDocsApp(
    initialDocumentUri: Uri? = null,
    initialDocumentType: DocumentType? = null,
    isOpenedExternally: Boolean = false,
) {
    val navController = rememberNavController()
    var selectedDestination by rememberSaveable { mutableStateOf(TopLevelDestination.FILES) }
    val activity = LocalContext.current as? Activity

    // Handle intent-based document opening
    LaunchedEffect(initialDocumentUri) {
        if (initialDocumentUri != null) {
            val docType = initialDocumentType ?: DocumentType.PDF
            navController.navigate(Routes.viewerForType(initialDocumentUri, docType)) {
                if (isOpenedExternally) {
                    popUpTo(navController.graph.startDestinationId) { inclusive = true }
                }
                launchSingleTop = true
            }
        }
    }

    NavHost(
        navController = navController,
        startDestination = TopLevelDestination.FILES.route,
        modifier = Modifier.fillMaxSize(),
    ) {
        // Main tabs with bottom nav
        composable(TopLevelDestination.FILES.route) {
            MainScaffold(
                selectedDestination = TopLevelDestination.FILES,
                onDestinationSelected = { dest ->
                    selectedDestination = dest
                    navController.navigate(dest.route) {
                        popUpTo(navController.graph.startDestinationId) {
                            saveState = true
                        }
                        launchSingleTop = true
                        restoreState = true
                    }
                },
            ) {
                HomeScreen(
                    onDocumentOpened = { uri, documentType, displayName ->
                        navController.navigate(Routes.viewerForType(uri, documentType, displayName))
                    },
                )
            }
        }

        composable(TopLevelDestination.RECENT.route) {
            MainScaffold(
                selectedDestination = TopLevelDestination.RECENT,
                onDestinationSelected = { dest ->
                    selectedDestination = dest
                    navController.navigate(dest.route) {
                        popUpTo(navController.graph.startDestinationId) {
                            saveState = true
                        }
                        launchSingleTop = true
                        restoreState = true
                    }
                },
            ) {
                PlaceholderScreen(title = "Recent Files")
            }
        }

        composable(TopLevelDestination.SETTINGS.route) {
            MainScaffold(
                selectedDestination = TopLevelDestination.SETTINGS,
                onDestinationSelected = { dest ->
                    selectedDestination = dest
                    navController.navigate(dest.route) {
                        popUpTo(navController.graph.startDestinationId) {
                            saveState = true
                        }
                        launchSingleTop = true
                        restoreState = true
                    }
                },
            ) {
                SettingsScreen()
            }
        }

        // PDF Viewer (full screen, no bottom nav)
        composable(
            route = Routes.PDF_VIEWER,
            arguments = listOf(
                navArgument("uri") { type = NavType.StringType },
                navArgument("name") {
                    type = NavType.StringType
                    defaultValue = ""
                },
            ),
        ) { backStackEntry ->
            val uriString = backStackEntry.arguments?.getString("uri") ?: ""
            val name = backStackEntry.arguments?.getString("name") ?: ""
            val uri = Uri.parse(uriString)
            val navigateBack = {
                if (!navController.popBackStack()) {
                    activity?.moveTaskToBack(true)
                }
                Unit
            }

            BackHandler(onBack = navigateBack)
            PdfViewerScreen(
                uri = uri,
                displayName = name.ifEmpty { null },
                onNavigateBack = navigateBack,
            )
        }

        // DOCX Viewer (full screen, no bottom nav)
        composable(
            route = Routes.DOCX_VIEWER,
            arguments = listOf(
                navArgument("uri") { type = NavType.StringType },
                navArgument("name") {
                    type = NavType.StringType
                    defaultValue = ""
                },
            ),
        ) { backStackEntry ->
            val uriString = backStackEntry.arguments?.getString("uri") ?: ""
            val name = backStackEntry.arguments?.getString("name") ?: ""
            val uri = Uri.parse(uriString)
            val navigateBack = {
                if (!navController.popBackStack()) {
                    activity?.moveTaskToBack(true)
                }
                Unit
            }

            BackHandler(onBack = navigateBack)
            DocxViewerScreen(
                uri = uri,
                displayName = name.ifEmpty { null },
                onNavigateBack = navigateBack,
            )
        }

        // XLSX Viewer (full screen, no bottom nav)
        composable(
            route = Routes.XLSX_VIEWER,
            arguments = listOf(
                navArgument("uri") { type = NavType.StringType },
                navArgument("name") {
                    type = NavType.StringType
                    defaultValue = ""
                },
            ),
        ) { backStackEntry ->
            val uriString = backStackEntry.arguments?.getString("uri") ?: ""
            val name = backStackEntry.arguments?.getString("name") ?: ""
            val uri = Uri.parse(uriString)
            val navigateBack = {
                if (!navController.popBackStack()) {
                    activity?.moveTaskToBack(true)
                }
                Unit
            }

            BackHandler(onBack = navigateBack)
            XlsxViewerScreen(
                uri = uri,
                displayName = name.ifEmpty { null },
                onNavigateBack = navigateBack,
            )
        }

        // PPTX Viewer (full screen, no bottom nav)
        composable(
            route = Routes.PPTX_VIEWER,
            arguments = listOf(
                navArgument("uri") { type = NavType.StringType },
                navArgument("name") {
                    type = NavType.StringType
                    defaultValue = ""
                },
            ),
        ) { backStackEntry ->
            val uriString = backStackEntry.arguments?.getString("uri") ?: ""
            val name = backStackEntry.arguments?.getString("name") ?: ""
            val uri = Uri.parse(uriString)
            val navigateBack = {
                if (!navController.popBackStack()) {
                    activity?.moveTaskToBack(true)
                }
                Unit
            }

            BackHandler(onBack = navigateBack)
            PptxViewerScreen(
                uri = uri,
                displayName = name.ifEmpty { null },
                onNavigateBack = navigateBack,
            )
        }
    }
}

@Composable
private fun MainScaffold(
    selectedDestination: TopLevelDestination,
    onDestinationSelected: (TopLevelDestination) -> Unit,
    content: @Composable () -> Unit,
) {
    NavigationSuiteScaffold(
        navigationSuiteItems = {
            TopLevelDestination.entries.forEach { destination ->
                item(
                    icon = {
                        Icon(
                            imageVector = if (destination == selectedDestination) {
                                destination.selectedIcon
                            } else {
                                destination.unselectedIcon
                            },
                            contentDescription = destination.label,
                        )
                    },
                    label = { Text(destination.label) },
                    selected = destination == selectedDestination,
                    onClick = { onDestinationSelected(destination) },
                )
            }
        },
    ) {
        content()
    }
}

@Composable
private fun PlaceholderScreen(title: String) {
    Scaffold { innerPadding ->
        Box(
            modifier = Modifier
                .fillMaxSize()
                .padding(innerPadding),
            contentAlignment = Alignment.Center,
        ) {
            Text(text = "$title — coming soon")
        }
    }
}
