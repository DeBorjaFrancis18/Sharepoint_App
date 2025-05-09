def upload_file_in_chunks(ctx, target_folder, file_path, file_name, chunk_size_mb=10):
    """Robust chunked upload with proper error handling"""
    chunk_size = chunk_size_mb * 1024 * 1024
    file_size = os.path.getsize(file_path)
    offset = 0  # Explicitly initialize offset
    upload_session = None
    
    try:
        print(f"Starting upload for '{file_name}' ({file_size/1024/1024:.2f} MB)")
        
        # 1. Create upload session
        upload_session = target_folder.files.create_upload_session(
            file_name, 
            file_size
        ).execute_query()
        print("Upload session created successfully")
        
        # 2. Upload chunks
        with open(file_path, 'rb') as f:
            while offset < file_size:
                chunk = f.read(chunk_size)
                if not chunk:
                    break
                
                is_last = (offset + len(chunk)) >= file_size
                
                # Upload with retry logic
                for attempt in range(3):
                    try:
                        if is_last:
                            uploaded_file = upload_session.finish_upload(offset, chunk).execute_query()
                        else:
                            upload_session.upload_chunk(offset, chunk).execute_query()
                        break
                    except Exception as e:
                        if attempt == 2:  # Final attempt
                            raise Exception(f"Failed to upload chunk at offset {offset}: {str(e)}")
                        print(f"Retrying chunk... (Attempt {attempt + 1})")
                        time.sleep(5)
                
                offset += len(chunk)
                print(f"Progress: {offset/1024/1024:.2f}MB/{file_size/1024/1024:.2f}MB")
                time.sleep(1)  # Avoid throttling
        
        # 3. Verification
        print("Verifying upload completion...")
        ctx.load(uploaded_file)
        ctx.execute_query()
        
        if uploaded_file.length != file_size:
            raise Exception(f"Size mismatch! Expected {file_size}, got {uploaded_file.length}")
        
        print(f"Successfully uploaded and verified '{file_name}'")
        return uploaded_file
        
    except Exception as e:
        error_msg = f"Upload failed at offset {offset if 'offset' in locals() else 'N/A'}: {str(e)}"
        print(error_msg)
        
        # Cleanup if upload failed
        if upload_session:
            try:
                print("Cleaning up failed upload session...")
                upload_session.delete_object().execute_query()
            except Exception as cleanup_error:
                print(f"Cleanup failed: {str(cleanup_error)}")
        
        # Check if file was partially uploaded
        existing_file = target_folder.files.get_by_name(file_name)
        try:
            ctx.load(existing_file)
            ctx.execute_query()
            actual_size = existing_file.properties.get('Length', 0)
            if actual_size > 0:
                print(f"Warning: Partial upload exists ({actual_size} bytes)")
        except:
            pass
        
        raise Exception(error_msg)
