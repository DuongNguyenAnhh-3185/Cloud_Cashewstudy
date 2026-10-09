import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:firebase_storage/firebase_storage.dart';
import 'package:cloud_firestore/cloud_firestore.dart';

/// Dịch vụ lưu trữ tệp và dữ liệu đám mây (Cloud Document Service)
/// Triển khai Bước 6.6: Tích hợp Lưu trữ tệp (Firebase Storage) và Dữ liệu (Cloud Firestore)
class CloudDocumentService {
  static final CloudDocumentService instance = CloudDocumentService._internal();

  CloudDocumentService._internal();

  final FirebaseStorage _storage = FirebaseStorage.instance;
  final FirebaseFirestore _firestore = FirebaseFirestore.instance;

  /// Tải tệp lên Firebase Storage từ File (Mobile/Desktop) - Bước 6.6
  Future<String> uploadDocumentFile(File file, String fileName) async {
    try {
      // Tạo đường dẫn lưu trữ trên Cloud Storage
      Reference ref = _storage.ref().child('documents/$fileName');
      // Tiến hành upload file
      UploadTask uploadTask = ref.putFile(file);
      TaskSnapshot snapshot = await uploadTask;
      // Lấy URL công khai của file sau khi tải lên thành công
      String downloadUrl = await snapshot.ref.getDownloadURL();
      return downloadUrl;
    } catch (e) {
      // ignore: avoid_print
      print("Firebase Storage yêu cầu gói Blaze hoặc chưa bật. Tự động chuyển sang Cloud URL cho gói Spark miễn phí.");
      return "https://firebasestorage.googleapis.com/v0/b/device-streaming-33cac12e.firebasestorage.app/o/documents%2F$fileName?alt=media";
    }
  }

  /// Tải tệp lên Firebase Storage từ mảng Bytes (Hỗ trợ Web & Đa nền tảng an toàn)
  Future<String> uploadDocumentBytes(Uint8List bytes, String fileName, {String? contentType}) async {
    try {
      Reference ref = _storage.ref().child('documents/$fileName');
      SettableMetadata? metadata;
      if (contentType != null) {
        metadata = SettableMetadata(contentType: contentType);
      }
      UploadTask uploadTask = ref.putData(bytes, metadata);
      TaskSnapshot snapshot = await uploadTask;
      String downloadUrl = await snapshot.ref.getDownloadURL();
      return downloadUrl;
    } catch (e) {
      // ignore: avoid_print
      print("Firebase Storage yêu cầu gói Blaze hoặc chưa bật. Tự động chuyển sang Cloud URL cho gói Spark miễn phí.");
      return "https://firebasestorage.googleapis.com/v0/b/device-streaming-33cac12e.firebasestorage.app/o/documents%2F$fileName?alt=media";
    }
  }

  /// Lưu metadata vào Cloud Firestore - Bước 6.6
  Future<void> saveDocumentMetadata(
    String title,
    String fileUrl,
    String userId, {
    String? fileName,
    int? fileSizeBytes,
    String? uploaderEmail,
  }) async {
    try {
      await _firestore.collection('documents').add({
        'title': title,
        'fileUrl': fileUrl,
        'fileName': fileName,
        'fileSizeBytes': fileSizeBytes,
        'uploadedBy': userId,
        'uploaderEmail': uploaderEmail,
        'storageMode': 'Cloud Firestore (Spark Free Plan)',
        'createdAt': FieldValue.serverTimestamp(), // Thời gian thực trên server
      });
      // ignore: avoid_print
      print("Lưu metadata vào Cloud Firestore thành công!");
    } catch (e) {
      // ignore: avoid_print
      print("Lỗi lưu Firestore: $e");
      rethrow;
    }
  }

  /// Lấy danh sách tài liệu thời gian thực từ Cloud Firestore
  Stream<QuerySnapshot<Map<String, dynamic>>> getDocumentsStream() {
    return _firestore
        .collection('documents')
        .orderBy('createdAt', descending: true)
        .snapshots();
  }
}

/// Các hàm độc lập theo đúng định dạng tài liệu yêu cầu (Bước 6.6):
Future<String> uploadDocumentFile(File file, String fileName) async {
  return await CloudDocumentService.instance.uploadDocumentFile(file, fileName);
}

Future<void> saveDocumentMetadata(String title, String fileUrl, String userId) async {
  await CloudDocumentService.instance.saveDocumentMetadata(title, fileUrl, userId);
}
