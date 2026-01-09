
import { adaptFileMetadata } from './metadataAdapter'
import { ISharePointClient, FieldDefinition, UserInfo } from '../types'
import { describe, it, expect, vi, beforeEach } from 'vitest'

// Mock ISharePointClient
const mockClient = {
    getListFields: vi.fn(),
    ensureUser: vi.fn(),
    searchTerm: vi.fn(),
} as unknown as ISharePointClient;

// Mock Fields
const mockFields: FieldDefinition[] = [
    { InternalName: 'Title', Title: 'Title', TypeAsString: 'Text', Hidden: false },
    { InternalName: 'TaxField', Title: 'My Taxonomy', TypeAsString: 'TaxonomyFieldType', Hidden: false, TermSetId: 'ts-guid-1' },
    { InternalName: 'TaxFieldMulti', Title: 'My Taxonomy Multi', TypeAsString: 'TaxonomyFieldTypeMulti', Hidden: false, TermSetId: 'ts-guid-1' },
    { InternalName: 'UserField', Title: 'Assigned To', TypeAsString: 'User', Hidden: false },
    { InternalName: 'MultiChoiceField', Title: 'Colors', TypeAsString: 'MultiChoice', Hidden: false },
    { InternalName: 'BoolField', Title: 'Active', TypeAsString: 'Boolean', Hidden: false },
    { InternalName: 'DateField', Title: 'Due Date', TypeAsString: 'DateTime', Hidden: false },
];

describe('metadataAdapter', () => {

    beforeEach(() => {
        vi.clearAllMocks();
    });

    it('should adapt basic text fields', async () => {
        (mockClient.getListFields as any).mockResolvedValue(mockFields);

        const payload = { Title: 'New Title' };
        const result = await adaptFileMetadata(mockClient, 'TestList', payload);

        expect(result).toEqual([
            { FieldName: 'Title', FieldValue: 'New Title' }
        ]);
    });

    it('should adapt taxonomy fields (object format)', async () => {
        (mockClient.getListFields as any).mockResolvedValue(mockFields);

        const payload = {
            TaxField: { Label: 'Tag1', TermGuid: 'guid-1' }
        };
        const result = await adaptFileMetadata(mockClient, 'TestList', payload);

        expect(result).toEqual([
            { FieldName: 'TaxField', FieldValue: 'Tag1|guid-1' }
        ]);
    });

    it('should adapt taxonomy fields (string label resolution)', async () => {
        (mockClient.getListFields as any).mockResolvedValue(mockFields);
        (mockClient.searchTerm as any).mockResolvedValue({ Label: 'Resolved Tag', TermGuid: 'resolved-guid-1' });

        const payload = { TaxField: 'Tag1' };
        const result = await adaptFileMetadata(mockClient, 'TestList', payload);

        expect(mockClient.searchTerm).toHaveBeenCalledWith('ts-guid-1', 'Tag1');
        expect(result).toEqual([
            { FieldName: 'TaxField', FieldValue: 'Resolved Tag|resolved-guid-1' }
        ]);
    });

    it('should adapt taxonomy multi fields (array of strings resolution)', async () => {
        (mockClient.getListFields as any).mockResolvedValue(mockFields);
        (mockClient.searchTerm as any)
            .mockResolvedValueOnce({ Label: 'Tag A', TermGuid: 'guid-a' })
            .mockResolvedValueOnce({ Label: 'Tag B', TermGuid: 'guid-b' });

        const payload = {
            TaxFieldMulti: ['Tag A', 'Tag B']
        };
        const result = await adaptFileMetadata(mockClient, 'TestList', payload);

        expect(mockClient.searchTerm).toHaveBeenCalledTimes(2);
        expect(result).toEqual([
            { FieldName: 'TaxFieldMulti', FieldValue: 'Tag A|guid-a;Tag B|guid-b' }
        ]);
    });

    it('should adapt user fields (email string)', async () => {
        (mockClient.getListFields as any).mockResolvedValue(mockFields);
        (mockClient.ensureUser as any).mockResolvedValue({ LoginName: 'i:0#.f|m|user@test.com' } as UserInfo);

        const payload = { UserField: 'user@test.com' };
        const result = await adaptFileMetadata(mockClient, 'TestList', payload);

        expect(result).toEqual([
            { FieldName: 'UserField', FieldValue: 'i:0#.f|m|user@test.com' }
        ]);
    });

    it('should adapt boolean fields', async () => {
        (mockClient.getListFields as any).mockResolvedValue(mockFields);

        const payload = { BoolField: true };
        const result = await adaptFileMetadata(mockClient, 'TestList', payload);

        expect(result).toEqual([
            { FieldName: 'BoolField', FieldValue: 'TRUE' }
        ]);

        const payload2 = { BoolField: false };
        const result2 = await adaptFileMetadata(mockClient, 'TestList', payload2);

        expect(result2).toEqual([
            { FieldName: 'BoolField', FieldValue: 'FALSE' }
        ]);
    });

    it('should adapt date fields', async () => {
        (mockClient.getListFields as any).mockResolvedValue(mockFields);
        const date = new Date('2023-01-01T12:00:00Z');

        const payload = { DateField: date };
        const result = await adaptFileMetadata(mockClient, 'TestList', payload);

        expect(result).toEqual([
            { FieldName: 'DateField', FieldValue: date.toISOString() }
        ]);
    });
});
