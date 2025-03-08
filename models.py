from sqlalchemy import Column, Integer, String, Float, DateTime, ForeignKey
from sqlalchemy.orm import relationship, declarative_base
from datetime import datetime

Base = declarative_base()

class Class(Base):  # type: ignore
    __tablename__ = 'classes'
    
    id = Column(Integer, primary_key=True)
    name = Column(String(255), nullable=False)
    created_at = Column(DateTime, nullable=False, default=datetime.utcnow)
    
    students = relationship("Student", back_populates="class_")

class Student(Base):  # type: ignore
    __tablename__ = 'students'
    
    id = Column(Integer, primary_key=True)
    class_id = Column(Integer, ForeignKey('classes.id'), nullable=False)
    student_id = Column(String(255), nullable=False)
    first_name = Column(String(255), nullable=False)
    last_name = Column(String(255), nullable=False)
    activities_grade = Column(Float, nullable=False)
    exam_grade = Column(Float, nullable=False)
    test_grade = Column(Float, nullable=False)
    final_grade = Column(Float, nullable=False)
    grade_comment = Column(String(255), nullable=True)
    created_at = Column(DateTime, nullable=False, default=datetime.utcnow)
    
    class_ = relationship("Class", back_populates="students")