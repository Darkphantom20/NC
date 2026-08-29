import { NextRequest, NextResponse } from 'next/server';
import { Packer } from 'docx';
import { buildReportDocument } from '@/lib/report-generator';

export async function POST(request: NextRequest) {
  try {
    const payload = await request.json();
    const doc = buildReportDocument(payload);
    const blob = await Packer.toBlob(doc);

    return new NextResponse(blob, {
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'Content-Disposition': 'attachment; filename="JRMSU_Narrative_Report_Sections_1_to_6.docx"',
      },
    });
  } catch (error) {
    console.error(error);
    return NextResponse.json({ error: 'Failed to generate document' }, { status: 500 });
  }
}
