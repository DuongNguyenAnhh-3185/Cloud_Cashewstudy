import 'package:flutter/material.dart';
import 'core/constants/app_constants.dart';
import 'presentation/screens/home_screen.dart';
import 'service_locator.dart';

void main() async {
  WidgetsFlutterBinding.ensureInitialized();

  // Khởi tạo Dependency Injection & CSDL SQLite Cục bộ
  await ServiceLocator.instance.initialize();

  runApp(const CashewDocApp());
}

class CashewDocApp extends StatelessWidget {
  const CashewDocApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: AppConstants.appName,
      debugShowCheckedModeBanner: false,
      theme: ThemeData(
        useMaterial3: true,
        colorScheme: ColorScheme.fromSeed(
          seedColor: AppConstants.primaryColor,
          primary: AppConstants.primaryColor,
          secondary: AppConstants.secondaryColor,
          surface: Colors.white,
        ),
        scaffoldBackgroundColor: AppConstants.backgroundColor,
        appBarTheme: const AppBarTheme(
          elevation: 0,
          backgroundColor: AppConstants.primaryColor,
          foregroundColor: Colors.white,
          centerTitle: false,
        ),
      ),
      home: const HomeScreen(),
    );
  }
}
