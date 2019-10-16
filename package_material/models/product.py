from functools import reduce

from sqlalchemy import TIMESTAMP, Column, ForeignKey, Integer, String, text, and_
from sqlalchemy.dialects.mysql import INTEGER
from sqlalchemy.orm import relationship, Query

from .base import Base, db_session
from .mixins import MetaMixin


class ProductQuery(Query):

    def search(self, keywords):
        if not keywords:
            return self

        criteria = []
        for keyword in keywords.split():
            keyword = '%' + keyword + '%'
            criteria.append(Product.internal_name.ilike(keyword))

        q = reduce(and_, criteria)

        return self.filter(q).distinct()


class Product(Base, MetaMixin):
    __tablename__ = 'products'
    __table_args__ = {
        "mysql_charset": "utf8"
    }

    query = db_session.query_property(ProductQuery)

    id = Column(Integer, primary_key=True, autoincrement=True)
    category_id = Column(INTEGER(unsigned=True), ForeignKey('categories.id'))
    internal_name = Column(String(64), unique=True)
    market_name = Column(String(64))
    part_a = Column(String(64), nullable=True)
    part_b = Column(String(64), nullable=True)
    ab_ratio = Column(String(16), nullable=True)
    color = Column(String(32), nullable=True)
    spec = Column(String(32), nullable=True)
    label_viscosity = Column(String(32), nullable=True)
    viscosity_width = Column(String(32), nullable=True)
    created_at = Column(TIMESTAMP(True), nullable=True, server_default=text('CURRENT_TIMESTAMP'))
    updated_at = Column(TIMESTAMP(True), nullable=True)

    category = relationship("Category", back_populates="products", lazy='joined')

    def to_dict(self):
        return {c.name: getattr(self, c.name) for c in self.__table__.columns}
