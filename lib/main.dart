import 'package:flutter/material.dart';
import 'package:firebase_core/firebase_core.dart';
import 'firebase_options.dart';
import 'core/constants/app_constants.dart';
import 'presentation/screens/home_screen.dart';
import 'service_locator.dart';

void main() async {
  WidgetsFlutterBinding.ensureInitialized();

  // 1. Khởi tạo Dependency Injection & CSDL SQLite Cục bộ (Giao diện gốc CashewDoc)
  await ServiceLocator.instance.initialize();

  // 2. Khởi tạo dịch vụ Firebase cho các tính năng đám mây
  try {
    await Firebase.initializeApp(
      options: DefaultFirebaseOptions.currentPlatform,
    );
  } catch (e) {
    // ignore: avoid_print
    print("Firebase init notice: $e");
  }

  runApp(const CashewDocApp());
}

/// Ứng dụng Quản lý Tài liệu Học tập CashewDoc (Giao diện ban đầu)
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

/// Bí danh giữ tương thích với các bài kiểm thử widget test
typedef MyApp = CashewDocApp;