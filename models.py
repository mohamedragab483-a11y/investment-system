from sqlalchemy import Column, Integer, String, Float, DateTime, ForeignKey, Text
from sqlalchemy.orm import relationship
from datetime import datetime
from database import Base

class User(Base):
    __tablename__ = "users"
    id = Column(Integer, primary_key=True, index=True)
    username = Column(String, unique=True, index=True)
    password_hash = Column(String)
    role = Column(String, default="admin")
    investor_id = Column(Integer, ForeignKey("investors.id"), nullable=True)
    investor = relationship("Investor")

class Investor(Base):
    __tablename__ = "investors"
    id = Column(Integer, primary_key=True, index=True)
    code = Column(String, unique=True, index=True)
    name = Column(String, index=True)
    phone = Column(String, nullable=True)
    notes = Column(Text, nullable=True)
    contributions = relationship("Contribution", back_populates="investor")

class Project(Base):
    __tablename__ = "projects"
    id = Column(Integer, primary_key=True, index=True)
    code = Column(String, unique=True, index=True)
    name = Column(String, index=True)
    project_type = Column(String, default="عقاري")
    status = Column(String, default="قيد التنفيذ")
    total_capital = Column(Float, default=0.0)
    contributions = relationship("Contribution", back_populates="project", cascade="all, delete-orphan")
    collections = relationship("Collection", back_populates="project", cascade="all, delete-orphan")
    expenses = relationship("Expense", back_populates="project", cascade="all, delete-orphan")
    sales = relationship("Sale", back_populates="project", cascade="all, delete-orphan")
    units = relationship("ProjectUnit", back_populates="project", cascade="all, delete-orphan")

class ProjectUnit(Base):
    __tablename__ = "project_units"
    id = Column(Integer, primary_key=True, index=True)
    project_id = Column(Integer, ForeignKey("projects.id"))
    unit_code = Column(String, index=True)
    unit_name = Column(String)
    unit_type = Column(String, default="شقة")
    area = Column(Float, default=0.0)
    cost_price = Column(Float, default=0.0)
    target_price = Column(Float, default=0.0)
    status = Column(String, default="متاحة")
    notes = Column(String, nullable=True)
    project = relationship("Project", back_populates="units")
    sales = relationship("Sale", back_populates="unit")

class Contribution(Base):
    __tablename__ = "contributions"
    id = Column(Integer, primary_key=True, index=True)
    project_id = Column(Integer, ForeignKey("projects.id"))
    investor_id = Column(Integer, ForeignKey("investors.id"))
    capital_share = Column(Float, default=0.0)
    mgmt_fee_pct = Column(Float, default=0.0)
    project = relationship("Project", back_populates="contributions")
    investor = relationship("Investor", back_populates="contributions")

class Collection(Base):
    __tablename__ = "collections"
    id = Column(Integer, primary_key=True, index=True)
    code = Column(String, unique=True, index=True)
    date = Column(DateTime, default=datetime.utcnow)
    project_id = Column(Integer, ForeignKey("projects.id"))
    amount = Column(Float, default=0.0)
    col_type = Column(String)
    description = Column(String, nullable=True)
    project = relationship("Project", back_populates="collections")

class Expense(Base):
    __tablename__ = "expenses"
    id = Column(Integer, primary_key=True, index=True)
    date = Column(DateTime, default=datetime.utcnow)
    project_id = Column(Integer, ForeignKey("projects.id"))
    exp_type = Column(String)
    description = Column(String, nullable=True)
    amount = Column(Float, default=0.0)
    project = relationship("Project", back_populates="expenses")

class Distribution(Base):
    __tablename__ = "distributions"
    id = Column(Integer, primary_key=True, index=True)
    code = Column(String, unique=True, index=True)
    date = Column(DateTime, default=datetime.utcnow)
    op_type = Column(String)
    source_project_id = Column(Integer, ForeignKey("projects.id"), nullable=True)
    target_project_id = Column(Integer, ForeignKey("projects.id"), nullable=True)
    investor_id = Column(Integer, ForeignKey("investors.id"), nullable=True)
    amount = Column(Float, default=0.0)
    description = Column(String, nullable=True)
    source_project = relationship("Project", foreign_keys=[source_project_id])
    target_project = relationship("Project", foreign_keys=[target_project_id])
    investor = relationship("Investor", foreign_keys=[investor_id])

class Sale(Base):
    __tablename__ = "sales"
    id = Column(Integer, primary_key=True, index=True)
    sale_code = Column(String, unique=True, index=True)
    project_id = Column(Integer, ForeignKey("projects.id"))
    unit_id = Column(Integer, ForeignKey("project_units.id"), nullable=True)
    buyer_name = Column(String)
    buyer_phone = Column(String, nullable=True)
    unit_description = Column(String)
    sale_date = Column(DateTime, default=datetime.utcnow)
    cost_price = Column(Float, default=0.0)
    total_price = Column(Float, default=0.0)
    net_unit_profit = Column(Float, default=0.0)
    sale_type = Column(String, default="تقسيط")
    down_payment = Column(Float, default=0.0)
    remaining_amount = Column(Float, default=0.0)
    installments_count = Column(Integer, default=0)
    period_type = Column(String, default="شهري")
    project = relationship("Project", back_populates="sales")
    unit = relationship("ProjectUnit", back_populates="sales")
    installments = relationship("Installment", back_populates="sale", cascade="all, delete-orphan")

class Installment(Base):
    __tablename__ = "installments"
    id = Column(Integer, primary_key=True, index=True)
    sale_id = Column(Integer, ForeignKey("sales.id"))
    project_id = Column(Integer, ForeignKey("projects.id"))
    installment_no = Column(Integer)
    due_date = Column(DateTime)
    amount = Column(Float, default=0.0)
    paid_amount = Column(Float, default=0.0)
    status = Column(String, default="مستحق")
    paid_date = Column(DateTime, nullable=True)
    notes = Column(String, nullable=True)
    sale = relationship("Sale", back_populates="installments")