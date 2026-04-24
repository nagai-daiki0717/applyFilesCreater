package app;

import software.amazon.awssdk.core.ResponseBytes;
import software.amazon.awssdk.core.sync.RequestBody;
import software.amazon.awssdk.regions.Region;
import software.amazon.awssdk.services.s3.S3Client;
import software.amazon.awssdk.services.s3.model.GetObjectRequest;
import software.amazon.awssdk.services.s3.model.PutObjectRequest;
import software.amazon.awssdk.services.s3.presigner.S3Presigner;
import software.amazon.awssdk.services.s3.presigner.model.GetObjectPresignRequest;
import software.amazon.awssdk.services.s3.presigner.model.PutObjectPresignRequest;
import software.amazon.awssdk.services.s3.presigner.model.PresignedGetObjectRequest;
import software.amazon.awssdk.services.s3.presigner.model.PresignedPutObjectRequest;
import software.amazon.awssdk.services.s3.model.DeleteObjectRequest;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;

public class S3Service {

    private final S3Client s3;

    private final S3Presigner presigner;

    public S3Service() {
        Region region = Region.of(Env.REGION);
        this.s3 = S3Client.builder().region(region).build();
        this.presigner = S3Presigner.builder().region(region).build();
    }

    public void downloadToFile(String bucket, String key,
            Path localPath) throws IOException {
        ResponseBytes<?> bytes = s3.getObjectAsBytes(GetObjectRequest.builder()
                .bucket(bucket).key(key).build());
        Files.write(localPath, bytes.asByteArray());
    }

    public byte[] downloadBytes(String bucket, String key) {
        return s3.getObjectAsBytes(GetObjectRequest.builder().bucket(bucket)
                .key(key).build()).asByteArray();
    }

    public void uploadFile(Path localPath, String bucket, String key,
            String contentType) {
        s3.putObject(PutObjectRequest.builder().bucket(bucket).key(key)
                .contentType(contentType).build(), RequestBody.fromFile(
                        localPath));
    }

    public String presignPutUrl(String bucket, String key, String contentType,
            long seconds) {
        PutObjectRequest request = PutObjectRequest.builder().bucket(bucket)
                .key(key).contentType(contentType).build();

        PresignedPutObjectRequest presigned = presigner.presignPutObject(
                PutObjectPresignRequest.builder().signatureDuration(Duration
                        .ofSeconds(seconds)).putObjectRequest(request).build());
        return presigned.url().toString();
    }

    public String presignGetUrl(String bucket, String key,
            String contentDisposition, String contentType, long seconds) {
        GetObjectRequest request = GetObjectRequest.builder().bucket(bucket)
                .key(key).responseContentDisposition(contentDisposition)
                .responseContentType(contentType).build();

        PresignedGetObjectRequest presigned = presigner.presignGetObject(
                GetObjectPresignRequest.builder().signatureDuration(Duration
                        .ofSeconds(seconds)).getObjectRequest(request).build());
        return presigned.url().toString();
    }

    public void deleteObject(String bucket, String key) {
        s3.deleteObject(DeleteObjectRequest.builder().bucket(bucket).key(key)
                .build());
    }
}
