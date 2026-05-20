"""API request/response models."""

from datetime import datetime
from enum import Enum
from typing import Optional

from pydantic import BaseModel, Field


class FormatResponse(BaseModel):
    """Response model for successful formatting."""

    success: bool = Field(..., description="Whether formatting succeeded")
    file_name: str = Field(..., description="Name of the output file")
    content_type: str = Field("application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    warnings: list[str] = Field(default_factory=list, description="Non-fatal warnings")
    processing_time_ms: Optional[int] = Field(None, description="Processing time in milliseconds")


class ErrorResponse(BaseModel):
    """Standard error response."""

    detail: str
    code: Optional[str] = None
    timestamp: datetime = Field(default_factory=datetime.utcnow)


class JobStatus(str, Enum):
    """Async job status."""

    PENDING = "pending"
    PROCESSING = "processing"
    COMPLETED = "completed"
    FAILED = "failed"


class JobResponse(BaseModel):
    """Async job status response."""

    job_id: str
    status: JobStatus
    result_url: Optional[str] = None
    error: Optional[str] = None
    created_at: datetime = Field(default_factory=datetime.utcnow)
    updated_at: Optional[datetime] = None
