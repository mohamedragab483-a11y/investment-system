from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

# ضع رابط نيون المنسوخ هنا بين علامتي التنصيص
SQLALCHEMY_DATABASE_URL = "postgresql://neondb_owner:npg_1ukfMWEdxoX8@ep-little-truth-b5sklstk-pooler.c-7.us-east-2.aws.neon.tech/neondb?sslmode=require&channel_binding=require"

# ملاحظة: تم إزالة check_same_thread لأنها خاصة بـ SQLite فقط
engine = create_engine(SQLALCHEMY_DATABASE_URL)

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

Base = declarative_base()


def get_db():
  db = SessionLocal()
  try:
    yield db
  finally:
    db.close()