-- Kiểm thử RLS bằng 3 tài khoản Firebase giả lập.
begin;
create extension if not exists pgtap with schema extensions;
select plan(7);

insert into public.users (id, email, display_name) values
  ('firebase-alice', 'alice@example.com', 'Alice'),
  ('firebase-bob', 'bob@example.com', 'Bob'),
  ('firebase-carol', 'carol@example.com', 'Carol');

insert into public.documents (
  id, owner_id, title, file_name, storage_path, file_size_bytes, mime_type
) values
  ('11111111-1111-4111-8111-111111111111', 'firebase-alice', 'Tài liệu Alice',
   'alice.pdf', 'documents/firebase-alice/doc-1/alice.pdf', 1024, 'application/pdf'),
  ('22222222-2222-4222-8222-222222222222', 'firebase-bob', 'Tài liệu Bob',
   'bob.pdf', 'documents/firebase-bob/doc-2/bob.pdf', 2048, 'application/pdf');

insert into public.permissions (document_id, user_id, access_level, granted_by)
values (
  '11111111-1111-4111-8111-111111111111',
  'firebase-bob',
  'editor',
  'firebase-alice'
);

-- Alice: chỉ nhìn thấy tài liệu do mình sở hữu.
select set_config(
  'request.jwt.claims',
  '{"sub":"firebase-alice","iss":"https://securetoken.google.com/device-streaming-33cac12e","aud":"device-streaming-33cac12e"}',
  true
);
set local role anon;
select extensions.is(
  (select count(*) from public.documents),
  1::bigint,
  'Alice chỉ thấy tài liệu của mình'
);
reset role;

-- Bob: thấy tài liệu riêng và tài liệu Alice chia sẻ với quyền editor.
select set_config(
  'request.jwt.claims',
  '{"sub":"firebase-bob","iss":"https://securetoken.google.com/device-streaming-33cac12e","aud":"device-streaming-33cac12e"}',
  true
);
set local role anon;
select extensions.is(
  (select count(*) from public.documents),
  2::bigint,
  'Bob thấy tài liệu riêng và tài liệu được chia sẻ'
);
select extensions.results_eq(
  $$
    update public.documents
    set title = 'Bob đã cập nhật'
    where id = '11111111-1111-4111-8111-111111111111'
    returning title
  $$,
  $$ values ('Bob đã cập nhật'::text) $$,
  'Bob editor cập nhật được metadata'
);
select extensions.is_empty(
  $$
    delete from public.documents
    where id = '11111111-1111-4111-8111-111111111111'
    returning id
  $$,
  'Bob không thể xóa tài liệu của Alice'
);
reset role;

-- Carol: trước khi được chia sẻ không nhìn thấy gì.
select set_config(
  'request.jwt.claims',
  '{"sub":"firebase-carol","iss":"https://securetoken.google.com/device-streaming-33cac12e","aud":"device-streaming-33cac12e"}',
  true
);
set local role anon;
select extensions.is(
  (select count(*) from public.documents),
  0::bigint,
  'Carol chưa được chia sẻ nên không thấy tài liệu'
);
reset role;

insert into public.permissions (document_id, user_id, access_level, granted_by)
values (
  '11111111-1111-4111-8111-111111111111',
  'firebase-carol',
  'viewer',
  'firebase-alice'
);

-- Carol: sau khi được chia sẻ viewer có thể đọc nhưng không thể sửa.
select set_config(
  'request.jwt.claims',
  '{"sub":"firebase-carol","iss":"https://securetoken.google.com/device-streaming-33cac12e","aud":"device-streaming-33cac12e"}',
  true
);
set local role anon;
select extensions.is(
  (select count(*) from public.documents),
  1::bigint,
  'Carol viewer thấy đúng tài liệu được chia sẻ'
);
select extensions.is_empty(
  $$
    update public.documents
    set title = 'Carol không được phép sửa'
    where id = '11111111-1111-4111-8111-111111111111'
    returning id
  $$,
  'Carol viewer không thể cập nhật metadata'
);
select extensions.is_empty(
  $$
    delete from public.documents
    where id = '11111111-1111-4111-8111-111111111111'
    returning id
  $$,
  'Carol viewer không thể xóa tài liệu'
);
reset role;

select * from extensions.finish();
rollback;
