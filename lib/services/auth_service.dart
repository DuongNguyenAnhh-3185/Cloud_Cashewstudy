import 'package:flutter/foundation.dart';
import 'package:firebase_auth/firebase_auth.dart';
import 'package:google_sign_in/google_sign_in.dart';

/// Dịch vụ xác thực người dùng (Authentication Service)
/// Triển khai Bước 6.5: Tích hợp Đăng nhập với Google (Google Sign-In)
class AuthService {
  static final AuthService instance = AuthService._internal();

  AuthService._internal();

  final FirebaseAuth _auth = FirebaseAuth.instance;
  GoogleSignIn? _googleSignIn;

  GoogleSignIn get _googleSignInInstance => _googleSignIn ??= GoogleSignIn();

  /// Lấy thông tin người dùng hiện tại đang đăng nhập
  User? get currentUser => _auth.currentUser;

  /// Stream theo dõi trạng thái đăng nhập thời gian thực
  Stream<User?> get authStateChanges => _auth.authStateChanges();

  /// Hàm xử lý đăng nhập bằng Google trong ứng dụng Flutter (Bước 6.5)
  Future<User?> signInWithGoogle() async {
    try {
      // 1. Trên nền tảng Web: Sử dụng chuẩn signInWithPopup của Firebase Web
      if (kIsWeb) {
        final GoogleAuthProvider googleProvider = GoogleAuthProvider();
        googleProvider.addScope('email');
        googleProvider.addScope('profile');
        final UserCredential userCredential = await _auth.signInWithPopup(googleProvider);
        return userCredential.user;
      }

      // 2. Trên Mobile / Desktop: Kích hoạt bảng chọn tài khoản Google
      final GoogleSignInAccount? googleUser = await _googleSignInInstance.signIn();
      if (googleUser == null) return null; // Người dùng hủy đăng nhập

      // Lấy thông tin xác thực (token) từ Google
      final GoogleSignInAuthentication googleAuth = await googleUser.authentication;

      // Tạo credential mới cho Firebase
      final credential = GoogleAuthProvider.credential(
        accessToken: googleAuth.accessToken,
        idToken: googleAuth.idToken,
      );

      // Xác thực với Firebase và trả về thông tin user
      UserCredential userCredential = await _auth.signInWithCredential(credential);
      return userCredential.user;
    } catch (e) {
      // ignore: avoid_print
      print("Lỗi đăng nhập Google: $e");
      return null;
    }
  }

  /// Đăng xuất khỏi Google và Firebase
  Future<void> signOut() async {
    try {
      if (!kIsWeb) {
        await _googleSignInInstance.signOut();
      }
      await _auth.signOut();
    } catch (e) {
      // ignore: avoid_print
      print("Lỗi đăng xuất: $e");
    }
  }
}

/// Hàm độc lập signInWithGoogle() theo đúng định dạng tài liệu yêu cầu (Bước 6.5)
Future<User?> signInWithGoogle() async {
  return await AuthService.instance.signInWithGoogle();
}
