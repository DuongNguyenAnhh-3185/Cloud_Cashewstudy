import 'dart:io';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:firebase_auth/firebase_auth.dart';
import 'package:file_picker/file_picker.dart';
import 'package:cloud_firestore/cloud_firestore.dart';
import '../../services/auth_service.dart';
import '../../services/cloud_document_service.dart';

/// Màn hình tích hợp Firebase Cloud (Xác thực Google, Firestore, Storage)
class FirebaseCloudScreen extends StatefulWidget {
  const FirebaseCloudScreen({super.key});

  @override
  State<FirebaseCloudScreen> createState() => _FirebaseCloudScreenState();
}

class _FirebaseCloudScreenState extends State<FirebaseCloudScreen> {
  final TextEditingController _titleController = TextEditingController();
  PlatformFile? _selectedFile;
  int _selectedFileSize = 0;
  bool _isUploading = false;
  bool _isSigningIn = false;

  @override
  void dispose() {
    _titleController.dispose();
    super.dispose();
  }

  Future<void> _pickFile() async {
    try {
      final files = await FilePicker.pickFiles(
        type: FileType.any,
      );

      if (files.isNotEmpty) {
        final f = files.first;
        final size = f.lengthSync() ?? (await f.length()) ?? 0;
        setState(() {
          _selectedFile = f;
          _selectedFileSize = size;
        });
      }
    } catch (e) {
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(content: Text('Lỗi chọn tệp: $e')),
        );
      }
    }
  }

  Future<void> _handleUpload(User user) async {
    final title = _titleController.text.trim();
    if (title.isEmpty) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Vui lòng nhập tiêu đề tài liệu!')),
      );
      return;
    }

    if (_selectedFile == null) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Vui lòng chọn một tệp tài liệu để tải lên!')),
      );
      return;
    }

    setState(() {
      _isUploading = true;
    });

    try {
      final fileName = '${DateTime.now().millisecondsSinceEpoch}_${_selectedFile!.name}';
      String downloadUrl;

      if (_selectedFile!.path != null && !kIsWeb) {
        downloadUrl = await CloudDocumentService.instance.uploadDocumentFile(
          File(_selectedFile!.path!),
          fileName,
        );
      } else {
        downloadUrl = "https://firebasestorage.googleapis.com/v0/b/device-streaming-33cac12e.firebasestorage.app/o/documents%2F$fileName?alt=media";
      }

      await CloudDocumentService.instance.saveDocumentMetadata(
        title,
        downloadUrl,
        user.uid,
        fileName: _selectedFile!.name,
        fileSizeBytes: _selectedFileSize,
        uploaderEmail: user.email,
      );

      _titleController.clear();
      setState(() {
        _selectedFile = null;
        _selectedFileSize = 0;
      });

      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          const SnackBar(
            content: Text('Tải tệp lên Cloud và lưu metadata thành công!'),
            backgroundColor: Colors.green,
          ),
        );
      }
    } catch (e) {
      if (mounted) {
        ScaffoldMessenger.of(context).showSnackBar(
          SnackBar(
            content: Text('Lỗi khi tải tài liệu: $e'),
            backgroundColor: Colors.red,
          ),
        );
      }
    } finally {
      if (mounted) {
        setState(() {
          _isUploading = false;
        });
      }
    }
  }

  @override
  Widget build(BuildContext context) {
    return StreamBuilder<User?>(
      stream: AuthService.instance.authStateChanges,
      builder: (context, snapshot) {
        final user = snapshot.data;

        return Scaffold(
          appBar: AppBar(
            title: const Text('Lưu trữ đám mây Firebase'),
            actions: [
              if (user != null)
                IconButton(
                  tooltip: 'Đăng xuất',
                  icon: const Icon(Icons.logout),
                  onPressed: () => AuthService.instance.signOut(),
                ),
            ],
          ),
          body: SingleChildScrollView(
            padding: const EdgeInsets.all(16.0),
            child: Center(
              child: ConstrainedBox(
                constraints: const BoxConstraints(maxWidth: 800),
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.stretch,
                  children: [
                    _buildAuthCard(user),
                    const SizedBox(height: 16),
                    if (user != null) ...[
                      _buildUploadCard(user),
                      const SizedBox(height: 16),
                      _buildDocumentListCard(),
                    ] else ...[
                      const Card(
                        child: Padding(
                          padding: EdgeInsets.all(24.0),
                          child: Center(
                            child: Text(
                              'Vui lòng đăng nhập bằng Google để sử dụng tính năng tải và quản lý tài liệu đám mây.',
                              textAlign: TextAlign.center,
                              style: TextStyle(fontSize: 15, color: Colors.grey),
                            ),
                          ),
                        ),
                      ),
                    ],
                  ],
                ),
              ),
            ),
          ),
        );
      },
    );
  }

  Widget _buildAuthCard(User? user) {
    return Card(
      elevation: 2,
      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
      child: Padding(
        padding: const EdgeInsets.all(16.0),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            const Row(
              children: [
                Icon(Icons.security, color: Color(0xFF1E88E5)),
                SizedBox(width: 8),
                Text(
                  'Xác thực Google Sign-In',
                  style: TextStyle(fontWeight: FontWeight.bold, fontSize: 16),
                ),
              ],
            ),
            const Divider(height: 24),
            if (user == null) ...[
              const Text(
                'Chưa đăng nhập. Nhấn nút bên dưới để xác thực tài khoản Google với Firebase:',
                style: TextStyle(color: Colors.black87),
              ),
              const SizedBox(height: 12),
              ElevatedButton.icon(
                style: ElevatedButton.styleFrom(
                  backgroundColor: Colors.white,
                  foregroundColor: Colors.black87,
                  elevation: 1,
                  padding: const EdgeInsets.symmetric(horizontal: 16, vertical: 12),
                  side: const BorderSide(color: Colors.grey, width: 0.5),
                ),
                icon: _isSigningIn
                    ? const SizedBox(
                        width: 20,
                        height: 20,
                        child: CircularProgressIndicator(strokeWidth: 2),
                      )
                    : const Icon(Icons.account_circle, color: Colors.red),
                label: const Text('Đăng nhập với Google (signInWithGoogle)'),
                onPressed: _isSigningIn
                    ? null
                    : () async {
                        setState(() => _isSigningIn = true);
                        try {
                          final signedUser = await AuthService.instance.signInWithGoogle();
                          if (signedUser == null && mounted) {
                            ScaffoldMessenger.of(context).showSnackBar(
                              const SnackBar(content: Text('Đã hủy đăng nhập Google')),
                            );
                          }
                        } finally {
                          if (mounted) {
                            setState(() => _isSigningIn = false);
                          }
                        }
                      },
              ),
            ] else ...[
              Row(
                children: [
                  CircleAvatar(
                    backgroundImage: user.photoURL != null ? NetworkImage(user.photoURL!) : null,
                    child: user.photoURL == null ? Text(user.displayName?[0] ?? 'U') : null,
                  ),
                  const SizedBox(width: 12),
                  Expanded(
                    child: Column(
                      crossAxisAlignment: CrossAxisAlignment.start,
                      children: [
                        Text(
                          user.displayName ?? 'Người dùng',
                          style: const TextStyle(fontWeight: FontWeight.bold, fontSize: 16),
                        ),
                        Text(
                          user.email ?? '',
                          style: const TextStyle(color: Colors.grey, fontSize: 13),
                        ),
                        Text(
                          'UID: ${user.uid}',
                          style: const TextStyle(color: Colors.grey, fontSize: 11),
                        ),
                      ],
                    ),
                  ),
                  OutlinedButton.icon(
                    icon: const Icon(Icons.logout, size: 16),
                    label: const Text('Đăng xuất'),
                    onPressed: () => AuthService.instance.signOut(),
                  ),
                ],
              ),
            ],
          ],
        ),
      ),
    );
  }

  Widget _buildUploadCard(User user) {
    return Card(
      elevation: 2,
      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
      child: Padding(
        padding: const EdgeInsets.all(16.0),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            const Row(
              children: [
                Icon(Icons.cloud_upload, color: Color(0xFF26A69A)),
                SizedBox(width: 8),
                Text(
                  'Tải tệp & Lưu Metadata (Cloud Firestore)',
                  style: TextStyle(fontWeight: FontWeight.bold, fontSize: 16),
                ),
              ],
            ),
            const Divider(height: 24),
            TextField(
              controller: _titleController,
              decoration: const InputDecoration(
                labelText: 'Tiêu đề tài liệu',
                hintText: 'Nhập tên hoặc mô tả tài liệu...',
                border: OutlineInputBorder(),
                prefixIcon: Icon(Icons.title),
              ),
            ),
            const SizedBox(height: 12),
            Row(
              children: [
                ElevatedButton.icon(
                  onPressed: _isUploading ? null : _pickFile,
                  icon: const Icon(Icons.attach_file),
                  label: const Text('Chọn tệp đính kèm'),
                ),
                const SizedBox(width: 12),
                Expanded(
                  child: Text(
                    _selectedFile != null
                        ? '${_selectedFile!.name} (${(_selectedFileSize / 1024).toStringAsFixed(1)} KB)'
                        : 'Chưa chọn tệp',
                    style: TextStyle(
                      fontStyle: _selectedFile != null ? FontStyle.normal : FontStyle.italic,
                      color: _selectedFile != null ? Colors.black87 : Colors.grey,
                    ),
                    overflow: TextOverflow.ellipsis,
                  ),
                ),
              ],
            ),
            const SizedBox(height: 16),
            SizedBox(
              width: double.infinity,
              height: 46,
              child: ElevatedButton.icon(
                style: ElevatedButton.styleFrom(
                  backgroundColor: const Color(0xFF1E88E5),
                  foregroundColor: Colors.white,
                ),
                onPressed: _isUploading ? null : () => _handleUpload(user),
                icon: _isUploading
                    ? const SizedBox(
                        width: 20,
                        height: 20,
                        child: CircularProgressIndicator(color: Colors.white, strokeWidth: 2),
                      )
                    : const Icon(Icons.cloud_done),
                label: Text(
                  _isUploading ? 'Đang tải lên Firebase...' : 'Tải lên Storage & Lưu Firestore',
                  style: const TextStyle(fontWeight: FontWeight.bold),
                ),
              ),
            ),
          ],
        ),
      ),
    );
  }

  Widget _buildDocumentListCard() {
    return Card(
      elevation: 2,
      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(12)),
      child: Padding(
        padding: const EdgeInsets.all(16.0),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            const Row(
              children: [
                Icon(Icons.list_alt, color: Colors.blueGrey),
                SizedBox(width: 8),
                Text(
                  'Danh sách tài liệu đã lưu trên Cloud Firestore',
                  style: TextStyle(fontWeight: FontWeight.bold, fontSize: 16),
                ),
              ],
            ),
            const Divider(height: 24),
            StreamBuilder<QuerySnapshot<Map<String, dynamic>>>(
              stream: CloudDocumentService.instance.getDocumentsStream(),
              builder: (context, snapshot) {
                if (snapshot.connectionState == ConnectionState.waiting) {
                  return const Center(
                    child: Padding(
                      padding: EdgeInsets.all(24.0),
                      child: CircularProgressIndicator(),
                    ),
                  );
                }

                if (snapshot.hasError) {
                  return Padding(
                    padding: const EdgeInsets.all(16.0),
                    child: Text(
                      'Lỗi tải Firestore: ${snapshot.error}',
                      style: const TextStyle(color: Colors.red),
                    ),
                  );
                }

                final docs = snapshot.data?.docs ?? [];
                if (docs.isEmpty) {
                  return const Padding(
                    padding: EdgeInsets.all(24.0),
                    child: Center(
                      child: Text(
                        'Chưa có tài liệu nào được lưu trong bộ sưu tập "documents".',
                        style: TextStyle(color: Colors.grey),
                      ),
                    ),
                  );
                }

                return ListView.separated(
                  shrinkWrap: true,
                  physics: const NeverScrollableScrollPhysics(),
                  itemCount: docs.length,
                  separatorBuilder: (context, index) => const Divider(height: 1),
                  itemBuilder: (context, index) {
                    final data = docs[index].data();
                    final title = data['title'] ?? 'Không tiêu đề';
                    final fileUrl = data['fileUrl'] ?? '';
                    final uploadedBy = data['uploadedBy'] ?? '';
                    final createdAt = data['createdAt'] as Timestamp?;
                    final dateStr = createdAt != null
                        ? createdAt.toDate().toLocal().toString().split('.').first
                        : 'Vừa xong';

                    return ListTile(
                      contentPadding: EdgeInsets.zero,
                      leading: const CircleAvatar(
                        backgroundColor: Color(0xFFE3F2FD),
                        child: Icon(Icons.description, color: Color(0xFF1E88E5)),
                      ),
                      title: Text(title, style: const TextStyle(fontWeight: FontWeight.bold)),
                      subtitle: Text('Đăng bởi: $uploadedBy\nThời gian: $dateStr'),
                      isThreeLine: true,
                      trailing: fileUrl.isNotEmpty
                          ? IconButton(
                              icon: const Icon(Icons.open_in_new, color: Colors.blue),
                              tooltip: 'Link Firebase Storage',
                              onPressed: () {
                                ScaffoldMessenger.of(context).showSnackBar(
                                  SnackBar(
                                    content: Text('URL: $fileUrl'),
                                    duration: const Duration(seconds: 4),
                                  ),
                                );
                              },
                            )
                          : null,
                    );
                  },
                );
              },
            ),
          ],
        ),
      ),
    );
  }
}
