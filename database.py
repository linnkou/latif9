from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
import os

# Create database engine
DATABASE_URL = "sqlite:///grades.db"
engine = create_engine(DATABASE_URL)

# Create session factory
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()