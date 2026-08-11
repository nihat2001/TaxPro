import uuid
from datetime import date, datetime

from sqlalchemy import Date, DateTime, String, Text, func
from sqlalchemy.dialects.postgresql import UUID
from sqlalchemy.orm import Mapped, mapped_column

from app.core.database import Base


class Profile(Base):
    __tablename__ = "profiles"

    id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True), primary_key=True, default=uuid.uuid4
    )
    ad_soyad: Mapped[str] = mapped_column(String(255), nullable=False)
    voen: Mapped[str] = mapped_column(String(50), nullable=False)
    nomre: Mapped[str] = mapped_column(String(50), nullable=False)
    asan_id: Mapped[str] = mapped_column(String(50), nullable=False)
    bitme_tarixi: Mapped[date | None] = mapped_column(Date, nullable=True)
    qeyd: Mapped[str | None] = mapped_column(Text, nullable=True)
    pin1: Mapped[str] = mapped_column(Text, nullable=False)
    pin2: Mapped[str] = mapped_column(Text, nullable=False)
    puk: Mapped[str] = mapped_column(Text, nullable=False)
    reng: Mapped[str] = mapped_column(String(20), default="#3b82f6")
    ikon: Mapped[str] = mapped_column(String(50), default="user")
    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now()
    )
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), server_default=func.now(), onupdate=func.now()
    )
