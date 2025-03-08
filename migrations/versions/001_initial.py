from alembic import op
import sqlalchemy as sa

def upgrade():
    # Create classes table
    op.create_table(
        'classes',
        sa.Column('id', sa.Integer(), primary_key=True),
        sa.Column('name', sa.String(255), nullable=False),
        sa.Column('created_at', sa.DateTime(), nullable=False)
    )

    # Create students table
    op.create_table(
        'students',
        sa.Column('id', sa.Integer(), primary_key=True),
        sa.Column('class_id', sa.Integer(), sa.ForeignKey('classes.id'), nullable=False),
        sa.Column('student_id', sa.String(255), nullable=False),
        sa.Column('first_name', sa.String(255), nullable=False),
        sa.Column('last_name', sa.String(255), nullable=False),
        sa.Column('activities_grade', sa.Float(), nullable=False),
        sa.Column('exam_grade', sa.Float(), nullable=False),
        sa.Column('test_grade', sa.Float(), nullable=False),
        sa.Column('final_grade', sa.Float(), nullable=False),
        sa.Column('grade_comment', sa.String(255), nullable=True),
        sa.Column('created_at', sa.DateTime(), nullable=False)
    )

def downgrade():
    op.drop_table('students')
    op.drop_table('classes')