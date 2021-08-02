<?php

namespace Acme\Command;

use DateTime;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use RetailCrm\Api\Enum\ByIdentifier;
use RetailCrm\Api\Factory\SimpleClientFactory;
use RetailCrm\Api\Model\Entity\Customers\Customer;
use RetailCrm\Api\Model\Entity\Customers\CustomerTag;
use RetailCrm\Api\Model\Request\Customers\CustomersEditRequest;
use RetailCrm\Api\Model\Request\Customers\CustomersRequest;
use Symfony\Component\Console\Command\Command;
use Symfony\Component\Console\Input\InputArgument;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Output\OutputInterface;
use Throwable;

class SetTagCommand extends Command
{
    const TAG = 'mayorista';

    private $client;

    private $customersAlreadyHaveTag;

    private $notFoundCustomers;

    public function __construct(string $name = null)
    {
        parent::__construct($name);
        $this->customersAlreadyHaveTag = [];
        $this->notFoundCustomers = [];
        $this->client = SimpleClientFactory::createClient();
    }

    protected function configure()
    {
        $this
            ->setName('rcrm:settag')
            ->addArgument('tag', InputArgument::REQUIRED)
        ;
    }

    protected function execute(InputInterface $input, OutputInterface $output)
    {
        $spreadsheet = IOFactory::load("base datos clientes.xlsx");
        $sheet = $spreadsheet->getSheet(0);
        $lastRow = $sheet->getHighestRow();
        $lastColumn = $sheet->getHighestColumn();

        for ($row = 2; $row <= $lastRow; $row++) {
            $rowData = $sheet->rangeToArray('A' . $row . ':' . $lastColumn . $row,NULL,TRUE,FALSE);
            $rowData = array_shift($rowData);
            $phone = $rowData[5];
            if (!$phone) {
                continue;
            }
            if (is_string($phone)) {
                $phone = str_replace(' ', '', $phone);
            }
            $phone = (string)$phone;

            $customer = $this->findCustomerByPhone($phone);
            if ($customer) {
                $this->addTagForCustomer($customer);
            } else {
                $this->notFoundCustomers[] = $phone;
            }
        }

        $this->saveNotFoundToExcel();

        return 0;
    }

    private function saveNotFoundToExcel()
    {
        unlink('notFound.xlsx');
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        foreach ($this->notFoundCustomers as $key => $phone) {
            $sheet->setCellValue('A' . $key, $phone);
        }

        $writer = new Xlsx($spreadsheet);
        $writer->save('notFound.xlsx');
    }

    private function findCustomerByPhone(string $phone)
    {
        $request = new CustomersRequest();
        $request->filter->name = $phone;
        $response = $this->client->customers->list($request);
        $customers = $response->customers;
        return array_shift($customers);
    }

    private function addTagForCustomer(Customer $customer)
    {
        $customerEditRequest           = new CustomersEditRequest();
        $customerEditRequest->by       = ByIdentifier::ID;
        $customerEditRequest->site     = 'shop-ru';

        $check = $this->checkTag($customer->tags);

        if ($check) {
            $this->customersAlreadyHaveTag[] = $customer;
            return;
        }

        $tag = new CustomerTag();
        $tag->name = self::TAG;
        $customer->tags = array_merge($customer->tags, [$tag]);
        $customer->birthday = null;
        $customerEditRequest->customer = $customer;


        try {
            $response = $this->client->customers->edit($customer->id, $customerEditRequest);

        } catch (Throwable $e){
            dd($e);
        }
    }

    private function checkTag(array $tags): bool
    {
        foreach ($tags as $existingTag) {
            if ($existingTag->name == self::TAG) {
                return true;
            }
        }
        return false;
    }
}