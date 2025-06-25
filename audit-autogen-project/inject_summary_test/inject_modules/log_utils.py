def log_write_text(log_list, status, label, val_init, val_final, location_or_reason):
    if status == "success":
        log_list.append(f"✅ {label} 期初={val_init}, 期末={val_final} → {location_or_reason}")
    elif status == "skip":
        log_list.append(f"⚠️ 跳过：{label} {location_or_reason}")
    elif status == "error":
        log_list.append(f"❌ {label} {location_or_reason}")
