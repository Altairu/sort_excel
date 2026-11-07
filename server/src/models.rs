use serde::{Deserialize, Serialize};

/// 案件レコード
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Record {
    pub id: i64,
    pub status: String, // 進行中, リスケ, 日程未定, 見積, 完了, キャンセル
    pub acceptance_date: String, // 受諾日 (YYYY-MM-DD)
    pub loan_date: String, // 融資実行日 (YYYY-MM-DD)
    pub loan_date_note: String, // 融資実行日補足/メモ
    pub form_type: String, // 形態 (中古マンション, 新築戸建など)
    pub client_name: String, // お客様氏名
    pub property_name: String, // 物件
    pub request_content: String, // 依頼内容 (設定, 保存・設定など)
    pub settlement_date: String, // 金消日・面談日 (YYYY-MM-DD)
    pub settlement_date_note: String, // 金消日補足 (予備日など)
    pub settlement_place: String, // 金消場所・面談場所
    pub settlement_time: String, // 金消時間 (HH:MM)
    pub staff: String, // 担当
    pub confirmation: String, // 意思確認
    pub doc_check: String, // 申請書類チェック
    pub agency: String, // ＡＲＵＨＩ代理店
    pub broker: String, // 仲介
    pub jurisdiction: String, // 管轄
    pub opponent_scrivener: String, // 先方書士
    pub remarks: String, // 備考
    pub attendance_time: String, // 立会時間 (HH:MM)
    pub attendance_place: String, // 立会場所
    pub attendee: String, // 立会者
    pub same_day_application: String, // 当日申請 (オンライン, 書面など)
    pub execution_time: String, // 実行時間
    pub progress_application: String, // 申請書
    pub progress_required_docs: String, // 必要書類
    pub progress_transfer_docs: String, // 移転書類2
    pub progress_complete_docs: String, // 完了書類
    pub progress_receive_date: String, // 受取日
    pub progress_tokyo_complete: String, // 東京へ完了
    pub progress_payment_date: String, // 入金日
    pub is_deleted: i32, // 0: 有効, 1: 削除
    pub version: i64, // 楽観的ロック用バージョン
    pub created_at: String,
    pub updated_at: String,
}

/// 新規作成/更新用リクエスト
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct RecordInput {
    pub status: Option<String>,
    pub acceptance_date: Option<String>,
    pub loan_date: Option<String>,
    pub loan_date_note: Option<String>,
    pub form_type: Option<String>,
    pub client_name: Option<String>,
    pub property_name: Option<String>,
    pub request_content: Option<String>,
    pub settlement_date: Option<String>,
    pub settlement_date_note: Option<String>,
    pub settlement_place: Option<String>,
    pub settlement_time: Option<String>,
    pub staff: Option<String>,
    pub confirmation: Option<String>,
    pub doc_check: Option<String>,
    pub agency: Option<String>,
    pub broker: Option<String>,
    pub jurisdiction: Option<String>,
    pub opponent_scrivener: Option<String>,
    pub remarks: Option<String>,
    pub attendance_time: Option<String>,
    pub attendance_place: Option<String>,
    pub attendee: Option<String>,
    pub same_day_application: Option<String>,
    pub execution_time: Option<String>,
    pub progress_application: Option<String>,
    pub progress_required_docs: Option<String>,
    pub progress_transfer_docs: Option<String>,
    pub progress_complete_docs: Option<String>,
    pub progress_receive_date: Option<String>,
    pub progress_tokyo_complete: Option<String>,
    pub progress_payment_date: Option<String>,
    pub version: Option<i64>,
    pub actor_name: Option<String>,
}

/// ソートキー設定
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SortKeyConfig {
    pub key: String,
    pub ascending: bool,
}

/// アクセス・操作履歴ログ
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AuditLog {
    pub id: i64,
    pub created_at: String,
    pub client_ip: String,
    pub actor_name: String,
    pub action_type: String, // ACCESS, CREATE, UPDATE, DELETE, RESTORE, PURGE, EXPORT, CONFIG, IMPORT, BACKUP
    pub target_id: Option<i64>,
    pub summary: String,
    pub details: String,
}

/// アプリケーション全体設定
#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct AppConfig {
    pub staff_order: Vec<String>,
    pub sort_keys: Vec<SortKeyConfig>,
    pub form_type_options: Vec<String>,
    pub request_content_options: Vec<String>,
    pub same_day_options: Vec<String>,
    pub status_options: Vec<String>,
    pub export_dir: String,
    pub export_file_name: String,
    pub backup_retention_days: u32,
    pub auto_export_on_change: bool,
}

impl Default for AppConfig {
    fn default() -> Self {
        Self {
            staff_order: vec![
                "北條".into(),
                "正木".into(),
                "河内".into(),
                "上野".into(),
                "吉田".into(),
                "杉田".into(),
                "岩川".into(),
                "椙村".into(),
                "中谷".into(),
                "脇本".into(),
            ],
            sort_keys: vec![
                SortKeyConfig {
                    key: "識別".into(),
                    ascending: false,
                },
                SortKeyConfig {
                    key: "担当".into(),
                    ascending: true,
                },
                SortKeyConfig {
                    key: "時間".into(),
                    ascending: true,
                },
            ],
            form_type_options: vec![
                "新築戸建".into(),
                "中古マンション".into(),
                "中古戸建".into(),
                "分割（アパートローン）".into(),
                "分割（アパートローン）完成".into(),
                "新規（アパートローン）".into(),
                "Ｒ投資".into(),
                "借換".into(),
                "借換／リフォーム".into(),
                "フリーローン".into(),
                "土地つなぎ".into(),
                "移転".into(),
                "保存".into(),
                "売却抹消".into(),
            ],
            request_content_options: vec![
                "設定".into(),
                "保存・設定".into(),
                "抹消・設定".into(),
                "設定のみ".into(),
                "移転・設定".into(),
                "名変・設定".into(),
                "保存".into(),
                "移転".into(),
                "抹消".into(),
            ],
            same_day_options: vec!["オンライン".into(), "書面".into()],
            status_options: vec![
                "進行中".into(),
                "日程未定".into(),
                "リスケ".into(),
                "見積".into(),
                "完了".into(),
                "キャンセル".into(),
            ],
            export_dir: "./exports".into(),
            export_file_name: "受託確認票.xlsx".into(),
            backup_retention_days: 7,
            auto_export_on_change: true,
        }
    }
}
