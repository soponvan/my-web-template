from __future__ import annotations

from datetime import datetime
from pathlib import Path
import json
import secrets


def _ensure_trading_file(manager):
    base_dir = Path(__file__).resolve().parent
    data_dir = base_dir / "sheep_farm_data"
    trading_file = data_dir / "trading_records.json"
    data_dir.mkdir(parents=True, exist_ok=True)
    if not trading_file.exists():
        trading_file.write_text("{}", encoding="utf-8")
    manager.trading_records = manager.load_json(trading_file)
    return trading_file


def _install_trading_routes():
    try:
        import app as farm_app  # type: ignore
    except Exception:
        return

    app = getattr(farm_app, "app", None)
    if app is None:
        return

    rules = {rule.rule for rule in app.url_map.iter_rules()}
    if "/api/trading" in rules:
        return

    manager = getattr(farm_app, "manager", None)
    if manager is None:
        return

    trading_file = _ensure_trading_file(manager)

    def _record_sale(animal_id, animal_type, buyer_name, quantity, unit_price, notes=""):
        if not animal_id:
            return {"status": "error", "message": "กรุณาระบุรหัสสัตว์"}
        if animal_id not in manager.sheep_records:
            return {"status": "error", "message": "ไม่พบสัตว์ ID นี้"}

        sheep = manager.sheep_records[animal_id]
        if sheep.get("status") in ["ตาย", "หาย", "ขายแล้ว"]:
            return {"status": "error", "message": "สัตว์นี้ไม่พร้อมสำหรับการขาย"}

        qty = float(quantity or 1)
        price = float(unit_price or 0)
        total_amount = qty * price

        trade_id = f"trade_sale_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}_{secrets.token_hex(4)}"
        manager.trading_records[trade_id] = {
            "id": trade_id,
            "type": "sale",
            "animal_id": animal_id,
            "animal_type": animal_type or sheep.get("animal_type", "แกะ"),
            "buyer_name": buyer_name or "-",
            "quantity": qty,
            "unit_price": price,
            "total_amount": total_amount,
            "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "notes": notes or "",
            "status": "สำเร็จ",
        }

        sheep["status"] = "ขายแล้ว"
        sheep["sold_date"] = datetime.now().strftime("%Y-%m-%d")
        manager.save_json(farm_app.SHEEP_FILE, manager.sheep_records)
        manager.save_json(trading_file, manager.trading_records)
        return {"status": "success", "message": "บันทึกการขายสำเร็จ"}

    def _record_purchase(animal_id, animal_type, seller_name, quantity, unit_price, notes=""):
        if not animal_id:
            return {"status": "error", "message": "กรุณาระบุรหัสสัตว์"}
        if animal_id in manager.sheep_records:
            return {"status": "error", "message": "สัตว์ ID นี้มีอยู่แล้วในระบบ"}

        qty = float(quantity or 1)
        price = float(unit_price or 0)
        total_amount = qty * price

        trade_id = f"trade_purchase_{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}_{secrets.token_hex(4)}"
        manager.trading_records[trade_id] = {
            "id": trade_id,
            "type": "purchase",
            "animal_id": animal_id,
            "animal_type": animal_type or "แกะ",
            "seller_name": seller_name or "-",
            "quantity": qty,
            "unit_price": price,
            "total_amount": total_amount,
            "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "notes": notes or "",
            "status": "สำเร็จ",
        }

        manager.sheep_records[animal_id] = {
            "id": animal_id,
            "animal_type": animal_type or "แกะ",
            "gender": "ไม่ระบุ",
            "dob": datetime.now().strftime("%Y-%m-%d"),
            "breed": "ไม่ระบุ",
            "weight": 0,
            "status": "ปกติ",
            "registered_date": datetime.now().strftime("%Y-%m-%d"),
            "source": "purchase",
        }
        manager.save_json(farm_app.SHEEP_FILE, manager.sheep_records)
        manager.save_json(trading_file, manager.trading_records)
        return {"status": "success", "message": "บันทึกการซื้อสำเร็จ"}

    def _get_trade_history():
        records = []
        for key, item in manager.trading_records.items():
            records.append({"id": key, "data": item})
        return sorted(records, key=lambda item: item["data"].get("date", ""), reverse=True)

    def _get_trade_summary():
        total_sales = sum(float(item.get("total_amount", 0) or 0) for item in manager.trading_records.values() if item.get("type") == "sale")
        total_purchases = sum(float(item.get("total_amount", 0) or 0) for item in manager.trading_records.values() if item.get("type") == "purchase")
        return {
            "total_sales": total_sales,
            "total_purchases": total_purchases,
            "total_profit": total_sales - total_purchases,
            "transactions": len(manager.trading_records),
        }

    if not hasattr(manager, "record_sale"):
        manager.record_sale = _record_sale
    if not hasattr(manager, "record_purchase"):
        manager.record_purchase = _record_purchase
    if not hasattr(manager, "get_trade_history"):
        manager.get_trade_history = _get_trade_history
    if not hasattr(manager, "get_trade_summary"):
        manager.get_trade_summary = _get_trade_summary

    @app.route("/api/trading", methods=["GET"])
    @farm_app.login_required
    def get_trading():
        return farm_app.jsonify({
            "history": manager.get_trade_history(),
            "summary": manager.get_trade_summary(),
        })

    @app.route("/api/trading/sale", methods=["POST"])
    @farm_app.permission_required("write")
    def sale_animal():
        data = farm_app.request.json or {}
        result = manager.record_sale(
            data.get("animal_id"),
            data.get("animal_type"),
            data.get("buyer_name"),
            data.get("quantity", 1),
            data.get("unit_price"),
            data.get("notes", ""),
        )
        return farm_app.jsonify(result)

    @app.route("/api/trading/purchase", methods=["POST"])
    @farm_app.permission_required("write")
    def purchase_animal():
        data = farm_app.request.json or {}
        result = manager.record_purchase(
            data.get("animal_id"),
            data.get("animal_type"),
            data.get("seller_name"),
            data.get("quantity", 1),
            data.get("unit_price"),
            data.get("notes", ""),
        )
        return farm_app.jsonify(result)


_install_trading_routes()
