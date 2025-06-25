def log_write(log, status, subject, detail=""):
    print(f"📝 log_write触发: {status}, {subject}, {detail}")  # 临时调试
    if status == "success":
        log.append(f"✅ {subject} {detail}")
    elif status == "skip":
        log.append(f"⚠️ 跳过：{subject} {detail}")
    elif status == "error":
        log.append(f"❗ 错误：{subject} {detail}")
    else:
        log.append(f"{status}: {subject} {detail}")
