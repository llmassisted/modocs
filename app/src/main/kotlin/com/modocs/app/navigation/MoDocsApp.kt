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
import androidx.navigation.compose.currentBackStackEntryAsState
import com.modocs.core.ui.components.DocumentSwitch
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

/**
 * A request to open an externally-supplied document. [token] is unique per
 * incoming intent so the navigation effect re-fires even when the same URI is
 * reopened (value-equal URIs would otherwise be skipped).
 */
data class DocumentRequest(
    val uri: Uri,
    val type: DocumentType,
    val token: Int,
)

object Routes {
    const val FILES = "files"
    const val RECENT = "recent"
    const val SETTINGS = "settings"
    const val PDF_VIEWER = "pdf_viewer?uri={uri}&name={name}"
    const val DOCX_VIEWER = "docx_viewer?uri={uri}&name={name}"
    const val XLSX_VIEWER = "xlsx_viewer?uri={uri}&name={name}"
    const val PPTX_VIEWER = "pptx_viewer?uri={uri}&name={name}"
    const val CSV_VIEWER = "csv_viewer?uri={uri}&name={name}"

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

    fun csvViewer(uri: Uri, displayName: String? = null): String {
        val encoded = Uri.encode(uri.toString())
        val name = Uri.encode(displayName ?: "")
        return "csv_viewer?uri=$encoded&name=$name"
    }

    /**
     * Route to the viewer for [documentType], or null when there is no viewer.
     *
     * Returning null rather than guessing is deliberate: this used to fall back
     * to the PDF viewer, so an unrecognised file opened as a garbled PDF. The
     * `when` is exhaustive on purpose too — adding a document type should be a
     * compile error here, not a silent fallback.
     */
    fun viewerForType(uri: Uri, documentType: DocumentType, displayName: String? = null): String? {
        return when (documentType) {
            DocumentType.PDF -> pdfViewer(uri, displayName)
            DocumentType.DOCX -> docxViewer(uri, displayName)
            DocumentType.XLSX -> xlsxViewer(uri, displayName)
            DocumentType.PPTX -> pptxViewer(uri, displayName)
            DocumentType.CSV -> csvViewer(uri, displayName)
            DocumentType.UNKNOWN -> null
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
    documentRequest: DocumentRequest?,
    onConsumed: (DocumentRequest) -> Unit,
    onOpened: (Uri, String, DocumentType) -> Unit,
) {
    val navController = rememberNavController()
    var selectedDestination by rememberSaveable { mutableStateOf(TopLevelDestination.FILES) }
    val context = LocalContext.current
    val activity = context as? Activity

    val currentEntry by navController.currentBackStackEntryAsState()
    val switch = documentRequest?.let { request ->
        val route = requireNotNull(Routes.viewerForType(request.uri, request.type))
        DocumentSwitch(request.token, open = {
            onConsumed(request)
            navController.navigate(route) {
                // The graph remains even after an external open removed its start destination.
                // Replace its children, always creating a fresh viewer and ViewModel.
                popUpTo(navController.graph.id)
            }
        }, cancel = { onConsumed(request) })
    }
    val destination = currentEntry?.destination?.route
    LaunchedEffect(documentRequest?.token, destination) {
        if (destination != null && destination !in setOf(Routes.PDF_VIEWER, Routes.DOCX_VIEWER, Routes.XLSX_VIEWER, Routes.CSV_VIEWER)) {
            switch?.open?.invoke()
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
                        // HomeViewModel already surfaces "Unsupported file format"
                        // before emitting, so an unviewable type just goes nowhere.
                        Routes.viewerForType(uri, documentType, displayName)?.let {
                            navController.navigate(it)
                        }
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
                HomeScreen(recentOnly = true, onDocumentOpened = { uri, type, name ->
                    Routes.viewerForType(uri, type, name)?.let { navController.navigate(it) }
                })
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
                if (navController.previousBackStackEntry == null) activity?.finish()
                else navController.popBackStack()
                Unit
            }

            BackHandler(onBack = navigateBack)
            PdfViewerScreen(
                onOpened = { name -> onOpened(uri, name, DocumentType.PDF) },
                documentSwitch = switch.takeIf { currentEntry == backStackEntry },
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
                if (navController.previousBackStackEntry == null) activity?.finish()
                else navController.popBackStack()
                Unit
            }

            BackHandler(onBack = navigateBack)
            DocxViewerScreen(
                onOpened = { name -> onOpened(uri, name, DocumentType.DOCX) },
                documentSwitch = switch.takeIf { currentEntry == backStackEntry },
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
                if (navController.previousBackStackEntry == null) activity?.finish()
                else navController.popBackStack()
                Unit
            }

            BackHandler(onBack = navigateBack)
            XlsxViewerScreen(
                onOpened = { name -> onOpened(uri, name, DocumentType.XLSX) },
                documentSwitch = switch.takeIf { currentEntry == backStackEntry },
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
                if (navController.previousBackStackEntry == null) activity?.finish()
                else navController.popBackStack()
                Unit
            }

            BackHandler(onBack = navigateBack)
            PptxViewerScreen(
                onOpened = { name -> onOpened(uri, name, DocumentType.PPTX) },
                uri = uri,
                displayName = name.ifEmpty { null },
                onNavigateBack = navigateBack,
            )
        }

        // CSV Viewer: the spreadsheet viewer, read-only (full screen, no bottom nav)
        composable(
            route = Routes.CSV_VIEWER,
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
                if (navController.previousBackStackEntry == null) activity?.finish()
                else navController.popBackStack()
                Unit
            }

            BackHandler(onBack = navigateBack)
            XlsxViewerScreen(
                onOpened = { name -> onOpened(uri, name, DocumentType.CSV) },
                documentSwitch = switch.takeIf { currentEntry == backStackEntry },
                uri = uri,
                displayName = name.ifEmpty { null },
                onNavigateBack = navigateBack,
                isCsv = true,
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
