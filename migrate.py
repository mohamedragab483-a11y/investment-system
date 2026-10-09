import models
from database import Base, engine

# 1. إنشاء الجداول تلقائياً على سيرفر نيون
print("جاري إنشاء الجداول في قاعدة بيانات نيون السحابية...")
Base.metadata.create_all(bind=engine)
print("تم إنشاء الجداول بنجاح! قاعدة البيانات جاهزة الآن لحفظ كل شيء للأبد.")