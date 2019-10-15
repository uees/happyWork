from sqlalchemy import TIMESTAMP, Column, Integer, String, Text, text
from sqlalchemy.orm import relationship

from .base import Base
from .mixins import MetaMixin


class Category(Base, MetaMixin):
    __tablename__ = 'categories'

    __table_args__ = {
        "mysql_charset": "utf8"
    }

    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(String(64), unique=True)
    slug = Column(String(64))
    memo = Column(Text, nullable=True)
    created_at = Column(TIMESTAMP(True), nullable=True, server_default=text('CURRENT_TIMESTAMP'))
    updated_at = Column(TIMESTAMP(True), nullable=True)

    products = relationship("Product", back_populates="category")
